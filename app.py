# app.py
# ------------------------------------------------------------
# Streamlit-app til at overføre musikrapport-data til tv-appl-en.xlsx
# ------------------------------------------------------------
#
# Kør lokalt sådan her i terminalen:
#   pip install -r requirements.txt
#   streamlit run app.py
#
# Vigtigt:
# - Appen bruger openpyxl til at åbne skabelonen og skrive værdier i eksisterende celler.
# - Den indsætter IKKE nye rækker og ændrer IKKE kolonnebredder, layout eller formatering.
# - Hvis der er flere sange end skabelonen har tomme rækker til, udfyldes kun de rækker,
#   der allerede findes i skabelonen. Appen viser en advarsel i browseren.

import io
import re
import traceback
import unicodedata
import zipfile
from datetime import datetime, time, timedelta

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


# ------------------------------------------------------------
# 1. STREAMLIT-SIDEOPSÆTNING
# ------------------------------------------------------------

st.set_page_config(
    page_title="Musikrapport → TV application",
    page_icon="🎵",
    layout="centered",
)

st.title("🎵 Musikrapport → TV application")
st.markdown(
    """
Upload først NCB-skabelonen **tv-appl-en.xlsx** og derefter musikrapport-exporten.
Når kildefilen er uploadet, finder appen automatisk alle faner. Du kan enten vælge
én bestemt fane eller udtrække alle faner på én gang. Hvis du vælger alle faner,
får du en ZIP-fil med én færdig Excel-rapport pr. fane.

Appen skriver kun værdier ind i de relevante celler og bevarer skabelonens layout,
kolonnebredder og formatering.
"""
)


# ------------------------------------------------------------
# 2. SMÅ HJÆLPEFUNKTIONER
# ------------------------------------------------------------

def is_empty(value) -> bool:
    """
    Tjekker om en celleværdi er tom.

    Excel-filer kan indeholde None, tomme strenge eller strenge med kun mellemrum.
    Denne funktion gør resten af koden mere robust.
    """
    return value is None or (isinstance(value, str) and value.strip() == "")


def safe_text(value) -> str:
    """
    Konverterer en celleværdi til tekst på en sikker måde.

    Hvis værdien er None, returneres en tom streng.
    Hvis værdien er et decimaltal som 39.0, returneres "39" i stedet for "39.0".
    """
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def normalize_text(value) -> str:
    """
    Normaliserer tekst, så vi lettere kan sammenligne labels og kolonnenavne.

    Eksempel:
    "Production Company:" bliver til "production company".
    """
    text = safe_text(value).lower().strip()
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    text = text.rstrip(":").strip()
    return text


def looks_like_label(value) -> bool:
    """
    Forsøger at vurdere, om en celle ligner et label i kildefilen.

    Mange metadatafelter i musikrapporten står som:
    "Production Title:"
    Derfor stopper vi med at lede mod højre, hvis vi rammer et nyt label.
    """
    return isinstance(value, str) and value.strip().endswith(":")


def unique_keep_order(values):
    """
    Fjerner dubletter, men bevarer rækkefølgen.

    Det er vigtigt ved komponister/forfattere:
    ["Smith", "Jones", "Smith"] bliver til ["Smith", "Jones"].
    """
    result = []
    seen = set()

    for value in values:
        text = safe_text(value)
        if not text:
            continue

        key = text.casefold()
        if key not in seen:
            seen.add(key)
            result.append(text)

    return result


# ------------------------------------------------------------
# 3. FUNKTIONER TIL AT FINDE METADATA I KILDEFILEN
# ------------------------------------------------------------

def find_value_next_to_label(ws, label_candidates, max_offset=6):
    """
    Finder værdien ud for et bestemt label i en Excel-fane.

    Eksempel i kildefilen:
    A12 = "Network/Station:"
    B12 = "TV 2"

    Funktionen finder label-cellen og returnerer den første ikke-tomme celle til højre.
    """
    labels = {normalize_text(label) for label in label_candidates}

    for row in ws.iter_rows():
        for cell in row:
            if normalize_text(cell.value) in labels:
                # Kig mod højre i samme række
                for offset in range(1, max_offset + 1):
                    target = ws.cell(row=cell.row, column=cell.column + offset)

                    # Hvis vi rammer et nyt label, stopper vi.
                    # Så undgår vi at komme til at læse et andet felt.
                    if looks_like_label(target.value):
                        return None

                    if not is_empty(target.value):
                        return target.value

                return None

    return None


def find_cell_right_of_label(ws, label_candidates):
    """
    Finder cellen lige til højre for et label i skabelonen.

    Eksempel i skabelonen:
    A3 = "Production company"
    B3 = tom celle, som skal udfyldes

    Funktionen returnerer selve celle-objektet, så vi kan skrive en værdi i den.
    """
    labels = {normalize_text(label) for label in label_candidates}

    for row in ws.iter_rows():
        for cell in row:
            if normalize_text(cell.value) in labels:
                return ws.cell(row=cell.row, column=cell.column + 1)

    return None


def extract_metadata(source_ws) -> dict:
    """
    Udtrækker metadata fra den valgte kildefane.

    Mapping:
    - Production company = Production Company:
    - Production title   = Production Title:
    - Episode number     = Episode Number:
    - Broadcaster        = Network/Station:
    """
    return {
        "Production company": find_value_next_to_label(
            source_ws,
            ["Production Company:"]
        ),
        "Production title": find_value_next_to_label(
            source_ws,
            ["Production Title:"]
        ),
        "Episode number": find_value_next_to_label(
            source_ws,
            ["Episode Number:"]
        ),
        "Broadcaster": find_value_next_to_label(
            source_ws,
            ["Network/Station:"]
        ),
    }


# ------------------------------------------------------------
# 4. FUNKTIONER TIL AT FINDE HEADER-RÆKKER OG KOLONNER
# ------------------------------------------------------------

def find_header_row(ws, required_headers):
    """
    Finder den række, hvor bestemte kolonneoverskrifter findes.

    I kildefilen leder vi fx efter:
    "Seq #" og "Music Title".

    I skabelonen leder vi fx efter:
    "Song title", "Artist", "Min", "Sec", "Track type".
    """
    required = {normalize_text(header) for header in required_headers}

    for row in ws.iter_rows():
        header_map = {}

        for cell in row:
            if not is_empty(cell.value):
                header_map[normalize_text(cell.value)] = cell.column

        if required.issubset(set(header_map.keys())):
            return row[0].row, header_map

    raise ValueError(
        f"Kunne ikke finde header-rækken med disse kolonner: {required_headers}"
    )


def get_col(header_map, candidates, required=True):
    """
    Finder en kolonne ud fra én eller flere mulige kolonnenavne.

    Eksempel:
    get_col(header_map, ["Music Performer", "Associated performer"])
    """
    for candidate in candidates:
        key = normalize_text(candidate)
        if key in header_map:
            return header_map[key]

    if required:
        raise ValueError(f"Mangler kolonne i filen. Prøvede: {candidates}")

    return None


def get_col_contains(header_map, candidates, required=True):
    """
    Finder en kolonne, hvor kolonnenavnet indeholder en bestemt tekst.

    Det bruges især til skabelonen, hvor headeren fx hedder:
    "Composers (Last name). Separate with /"

    Her vil vi gerne finde den bare ved at søge på "Composers".
    """
    for candidate in candidates:
        needle = normalize_text(candidate)

        for header_name, col_index in header_map.items():
            if needle == header_name or needle in header_name:
                return col_index

    if required:
        raise ValueError(f"Mangler kolonne i skabelonen. Prøvede: {candidates}")

    return None


# ------------------------------------------------------------
# 5. NAVNE-RENSNING: KOMPONISTER OG FORFATTERE
# ------------------------------------------------------------

def clean_party_name(name) -> str:
    """
    Rydder op i navne fra "Name of Music Interested Party".

    Fjerner fx:
    - [PRS]
    - [BMI]
    - (PRS)
    - (BMI)

    Eksempel:
    "John Williams [BMI]" bliver til "John Williams".
    """
    text = safe_text(name)

    # Fjern alt i kantede parenteser: [PRS]
    text = re.sub(r"\[[^\]]*\]", "", text)

    # Fjern alt i almindelige parenteser: (BMI)
    text = re.sub(r"\([^)]*\)", "", text)

    # Ryd op i mellemrum og tegn
    text = text.replace(";", " ")
    text = re.sub(r"\s+", " ", text).strip(" ,;/")

    return text


def split_party_names(raw_name):
    """
    Splitter navne, hvis flere navne mod forventning står i samme celle.

    Normalt står der ét navn per række, men denne funktion gør appen mere robust.
    Vi splitter forsigtigt på / og ;.
    """
    cleaned = clean_party_name(raw_name)

    if not cleaned:
        return []

    parts = re.split(r"\s*/\s*|\s*;\s*", cleaned)
    return [part.strip() for part in parts if part.strip()]


def extract_last_name(name) -> str:
    """
    Udtrækker efternavnet fra et navn.

    Eksempler:
    - "A. Z. Børresen" bliver til "Børresen"
    - "John Williams" bliver til "Williams"
    - "Sting" bliver til "Sting"
    - "Williams, John" bliver til "Williams"
    """
    text = clean_party_name(name)

    if not text:
        return ""

    # Hvis navnet står som "Efternavn, Fornavn", tager vi delen før kommaet.
    if "," in text:
        return text.split(",", 1)[0].strip()

    parts = text.split()

    # Ét ord: fx "Sting" eller "EJAE"
    if len(parts) == 1:
        return parts[0]

    # Håndter navne med suffix, fx "Tom Jesso Jr."
    suffixes = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}
    last_token_normalized = parts[-1].lower().strip(".")

    if last_token_normalized in {suffix.strip(".") for suffix in suffixes} and len(parts) >= 2:
        return f"{parts[-2]} {parts[-1]}"

    # Standard: sidste ord er efternavnet
    return parts[-1]


# ------------------------------------------------------------
# 6. TIDS-PARSING: MUSIC DURATION → MIN/SEC
# ------------------------------------------------------------

def parse_duration_to_min_sec(value):
    """
    Konverterer "Music Duration" til minutter og sekunder.

    Funktionen forsøger at håndtere flere formater:
    - "00:02:15"
    - "02:15"
    - "0.02:26"
    - Excel time-værdier som decimaltal
    - datetime.time
    - timedelta
    """
    if is_empty(value):
        return None, None

    total_seconds = None

    # Hvis Excel/openpyxl giver os en timedelta
    if isinstance(value, timedelta):
        total_seconds = int(round(value.total_seconds()))

    # Hvis værdien er en datetime
    elif isinstance(value, datetime):
        total_seconds = value.hour * 3600 + value.minute * 60 + value.second

    # Hvis værdien er et time-objekt
    elif isinstance(value, time):
        total_seconds = value.hour * 3600 + value.minute * 60 + value.second

    # Hvis værdien er et tal
    elif isinstance(value, (int, float)):
        # Excel gemmer nogle gange tid som en brøkdel af et døgn.
        # 0.5 = 12 timer.
        if value < 1:
            total_seconds = int(round(value * 24 * 60 * 60))
        else:
            # Hvis tallet er større end 1, antager vi, at det er sekunder.
            total_seconds = int(round(value))

    # Hvis værdien er tekst
    else:
        text = safe_text(value).strip()

        # Nogle exports kan skrive 0.02:26 i stedet for 00:02:26.
        # Vi gør det læsbart ved at skifte punktum til kolon, hvis der også er kolon.
        text = text.replace(",", ":")
        if "." in text and ":" in text:
            text = text.replace(".", ":")

        # Fjern mellemrum og udtræk tal
        text = re.sub(r"\s+", "", text)
        numbers = re.findall(r"\d+", text)

        if len(numbers) >= 3:
            # Brug de sidste tre tal som timer, minutter, sekunder
            hours = int(numbers[-3])
            minutes = int(numbers[-2])
            seconds = int(numbers[-1])
            total_seconds = hours * 3600 + minutes * 60 + seconds

        elif len(numbers) == 2:
            # MM:SS
            minutes = int(numbers[0])
            seconds = int(numbers[1])
            total_seconds = minutes * 60 + seconds

        elif len(numbers) == 1:
            # Kun ét tal: antag sekunder
            total_seconds = int(numbers[0])

    if total_seconds is None:
        return None, None

    if total_seconds < 0:
        total_seconds = 0

    minutes = total_seconds // 60
    seconds = total_seconds % 60

    return int(minutes), int(seconds)


# ------------------------------------------------------------
# 7. TRACK TYPE-MAPPING
# ------------------------------------------------------------

TRACK_TYPE_MAP = {
    "pre-existing": "Existing music",
    "pre existing": "Existing music",
    "preexisting": "Existing music",
    "specially commissioned": "Commissioned music",
    "special commissioned": "Commissioned music",
    "commissioned": "Commissioned music",
    "library/production": "Library music",
    "library production": "Library music",
    "library": "Library music",
    "production music": "Library music",
}


def translate_track_type(value) -> str:
    """
    Oversætter "Music Source" fra kildefilen til skabelonens Track type.

    Krævet mapping:
    - Pre-existing → Existing music
    - Specially Commissioned → Commissioned music
    - Library/Production → Library music
    """
    text = normalize_text(value)

    if not text:
        return ""

    if text in TRACK_TYPE_MAP:
        return TRACK_TYPE_MAP[text]

    # Fallback: hvis teksten indeholder et kendt mønster
    for key, mapped_value in TRACK_TYPE_MAP.items():
        if key in text:
            return mapped_value

    # Hvis værdien er ukendt, returnerer vi den oprindelige tekst.
    # Så kan brugeren se, at der var noget, der ikke matchede standardmappingen.
    return safe_text(value)


# ------------------------------------------------------------
# 8. UDTRÆK MUSIKDATA FRA KILDEFILEN
# ------------------------------------------------------------

def extract_music_rows(source_ws) -> list[dict]:
    """
    Udtrækker alle sange fra den valgte kildefane.

    Logik:
    - Find header-rækken med "Seq #" og "Music Title"
    - Gruppér efter Seq #
    - Linjer under en sang uden Seq # tilhører seneste Seq #
    - Tomme rækker ignoreres
    """
    header_row, header_map = find_header_row(
        source_ws,
        required_headers=["Seq #", "Music Title"]
    )

    # Find de relevante kolonner i kildefilen
    col_seq = get_col(header_map, ["Seq #"])
    col_title = get_col(header_map, ["Music Title"])
    col_duration = get_col(header_map, ["Music Duration"])
    col_source = get_col(header_map, ["Music Source"])

    # Ikke alle fremtidige exports er nødvendigvis helt ens,
    # så disse tre gøres ikke hårdt påkrævede.
    col_performer = get_col(header_map, ["Music Performer"], required=False)
    col_role = get_col(header_map, ["Music Interested Party Role"], required=False)
    col_name = get_col(header_map, ["Name of Music Interested Party"], required=False)

    groups = []
    current_group = None

    # Gå igennem alle rækker under headeren
    for row_number in range(header_row + 1, source_ws.max_row + 1):
        row_values = [
            source_ws.cell(row=row_number, column=col).value
            for col in range(1, source_ws.max_column + 1)
        ]

        # Spring helt tomme rækker over
        if all(is_empty(value) for value in row_values):
            continue

        seq_value = source_ws.cell(row=row_number, column=col_seq).value

        row_data = {
            "seq": seq_value,
            "title": source_ws.cell(row=row_number, column=col_title).value,
            "duration": source_ws.cell(row=row_number, column=col_duration).value,
            "source": source_ws.cell(row=row_number, column=col_source).value,
            "performer": source_ws.cell(row=row_number, column=col_performer).value if col_performer else None,
            "role": source_ws.cell(row=row_number, column=col_role).value if col_role else None,
            "name": source_ws.cell(row=row_number, column=col_name).value if col_name else None,
        }

        # Hvis rækken har Seq #, starter en ny sang
        if not is_empty(seq_value):
            current_group = {
                "seq": seq_value,
                "rows": [row_data],
            }
            groups.append(current_group)

        # Hvis rækken ikke har Seq #, hører den til seneste sang
        elif current_group is not None:
            current_group["rows"].append(row_data)

    songs = []

    for group in groups:
        rows = group["rows"]

        # Find første ikke-tomme titel, performer, duration og source i gruppen
        title = next(
            (safe_text(row["title"]) for row in rows if not is_empty(row["title"])),
            ""
        )

        performer = next(
            (safe_text(row["performer"]) for row in rows if not is_empty(row["performer"])),
            ""
        )

        duration = next(
            (row["duration"] for row in rows if not is_empty(row["duration"])),
            None
        )

        source = next(
            (row["source"] for row in rows if not is_empty(row["source"])),
            None
        )

        # Hvis Music Performer ikke er udfyldt, prøver vi Associated performer-linjer
        if not performer:
            associated_performers = []

            for row in rows:
                role = normalize_text(row["role"])

                if "associated performer" in role and not is_empty(row["name"]):
                    associated_performers.append(clean_party_name(row["name"]))

            performer = "/".join(unique_keep_order(associated_performers))

        composers = []
        writers = []

        for row in rows:
            role = normalize_text(row["role"])
            raw_name = row["name"]

            if is_empty(role) or is_empty(raw_name):
                continue

            # Split i tilfælde af flere navne i samme celle
            for one_name in split_party_names(raw_name):
                last_name = extract_last_name(one_name)

                if not last_name:
                    continue

                # Composer og Composer/Author ryger i komponistfeltet
                if "composer" in role:
                    composers.append(last_name)

                # Author, Writer, Lyricist og Composer/Author ryger i writerfeltet
                if "author" in role or "writer" in role or "lyricist" in role:
                    writers.append(last_name)

        minutes, seconds = parse_duration_to_min_sec(duration)

        songs.append(
            {
                "seq": group["seq"],
                "Song title": title,
                "Artist": performer,
                "Composers": "/".join(unique_keep_order(composers)),
                "Writers": "/".join(unique_keep_order(writers)),
                "Min": minutes,
                "Sec": seconds,
                "Track type": translate_track_type(source),
            }
        )

    return songs


# ------------------------------------------------------------
# 9. UDFYLD SKABELONEN
# ------------------------------------------------------------

def find_template_music_columns(template_ws):
    """
    Finder kolonnerne i Music content-tabellen i skabelonen.

    Appen går ikke ud fra faste kolonnebogstaver.
    Den finder selv header-rækken og kolonnerne ud fra overskrifterne.
    """
    header_row, header_map = find_header_row(
        template_ws,
        required_headers=["Song title", "Artist", "Min", "Sec", "Track type"]
    )

    columns = {
        "Song title": get_col(header_map, ["Song title"]),
        "Artist": get_col(header_map, ["Artist"]),
        "Composers": get_col_contains(header_map, ["Composers"]),
        "Writers": get_col_contains(header_map, ["Writers"]),
        "Min": get_col(header_map, ["Min"]),
        "Sec": get_col(header_map, ["Sec"]),
        "Track type": get_col(header_map, ["Track type"]),
    }

    return header_row, columns


def write_metadata_to_template(template_ws, metadata: dict):
    """
    Skriver metadata ind i skabelonens øverste felter.

    Funktionen ændrer kun cellen til højre for label-feltet.
    """
    for target_label, value in metadata.items():
        target_cell = find_cell_right_of_label(template_ws, [target_label])

        if target_cell is None:
            raise ValueError(f"Kunne ikke finde metadatafeltet i skabelonen: {target_label}")

        # Vi skriver kun værdien. Formatering og layout bevares.
        target_cell.value = value


def write_music_to_template(template_ws, songs: list[dict]) -> dict:
    """
    Skriver musikdata ind i Music content-tabellen.

    Vigtigt:
    - Der indsættes ikke nye rækker.
    - Der ændres ikke formatering.
    - Der skrives kun i eksisterende celler i de fundne music-kolonner.
    """
    header_row, columns = find_template_music_columns(template_ws)
    first_data_row = header_row + 1

    # Antal eksisterende rækker i skabelonen, som vi må skrive i.
    # Vi bruger max_row, fordi det er den struktur, skabelonen allerede har.
    capacity = template_ws.max_row - first_data_row + 1

    warnings = []

    if capacity <= 0:
        raise ValueError("Skabelonen har ingen tomme rækker under Music content-headeren.")

    # Ryd kun gamle værdier i de relevante musik-kolonner.
    # Dette ændrer ikke formatering, kolonnebredder eller layout.
    for row_number in range(first_data_row, template_ws.max_row + 1):
        for col_index in columns.values():
            template_ws.cell(row=row_number, column=col_index).value = None

    songs_to_write = songs[:capacity]
    skipped_count = max(0, len(songs) - capacity)

    if skipped_count > 0:
        warnings.append(
            f"Der var {len(songs)} musiknumre i kilden, men skabelonen har kun "
            f"plads til {capacity} eksisterende rækker. "
            f"{skipped_count} musiknumre blev derfor ikke skrevet ind. "
            f"Appen indsætter ikke nye rækker, fordi det ville ændre skabelonens struktur."
        )

    # Skriv sangene ind i skabelonen
    for index, song in enumerate(songs_to_write):
        row_number = first_data_row + index

        template_ws.cell(row=row_number, column=columns["Song title"]).value = song["Song title"]
        template_ws.cell(row=row_number, column=columns["Artist"]).value = song["Artist"]
        template_ws.cell(row=row_number, column=columns["Composers"]).value = song["Composers"]
        template_ws.cell(row=row_number, column=columns["Writers"]).value = song["Writers"]
        template_ws.cell(row=row_number, column=columns["Min"]).value = song["Min"]
        template_ws.cell(row=row_number, column=columns["Sec"]).value = song["Sec"]
        template_ws.cell(row=row_number, column=columns["Track type"]).value = song["Track type"]

    return {
        "written_count": len(songs_to_write),
        "skipped_count": skipped_count,
        "capacity": capacity,
        "warnings": warnings,
    }


# ------------------------------------------------------------
# 10. HOVEDFUNKTION: FRA TO UPLOADEDE FILER TIL FÆRDIG EXCEL
# ------------------------------------------------------------

def get_sheet_names(source_file_bytes: bytes) -> list[str]:
    """
    Læser alle fanenavne fra kildefilen.

    Dette bruges til den dynamiske dropdown-menu.
    """
    workbook = load_workbook(
        io.BytesIO(source_file_bytes),
        read_only=True,
        data_only=True,
    )
    return workbook.sheetnames


def process_files(template_file_bytes: bytes, source_file_bytes: bytes, selected_sheet: str):
    """
    Hele databehandlingen samlet ét sted.

    Input:
    - template_file_bytes: bytes fra uploadet tv-appl-en.xlsx
    - source_file_bytes: bytes fra uploadet musikrapport-export
    - selected_sheet: fanen brugeren har valgt i dropdown

    Output:
    - output_bytes: færdig Excel-fil som bytes
    - summary: info til brugerfladen
    """
    # Åbn kildefilen
    source_wb = load_workbook(
        io.BytesIO(source_file_bytes),
        read_only=False,
        data_only=True,
    )

    if selected_sheet not in source_wb.sheetnames:
        raise ValueError(f"Den valgte fane findes ikke i kildefilen: {selected_sheet}")

    source_ws = source_wb[selected_sheet]

    # Udtræk metadata og musikdata
    metadata = extract_metadata(source_ws)
    songs = extract_music_rows(source_ws)

    # Åbn skabelonen
    template_wb = load_workbook(
        io.BytesIO(template_file_bytes),
        read_only=False,
        data_only=False,
    )

    # Brug Application-fanen, hvis den findes. Ellers brug aktiv fane.
    if "Application" in template_wb.sheetnames:
        template_ws = template_wb["Application"]
    else:
        template_ws = template_wb.active

    # Skriv data ind i skabelonen
    write_metadata_to_template(template_ws, metadata)
    write_result = write_music_to_template(template_ws, songs)

    # Gem workbook i hukommelsen i stedet for på disk.
    # Det gør, at Streamlit kan sende filen direkte til download-knappen.
    output = io.BytesIO()
    template_wb.save(output)
    output.seek(0)

    summary = {
        "selected_sheet": selected_sheet,
        "metadata": metadata,
        "song_count_found": len(songs),
        "song_count_written": write_result["written_count"],
        "song_count_skipped": write_result["skipped_count"],
        "template_capacity": write_result["capacity"],
        "warnings": write_result["warnings"],
    }

    return output.getvalue(), summary


def make_unique_zip_name(filename: str, used_names: set[str]) -> str:
    """
    Sørger for, at filnavne inde i ZIP-filen er unikke.

    Hvis to faner mod forventning giver samme filnavn, laver vi fx:
    - rapport.xlsx
    - rapport_2.xlsx
    """
    if filename not in used_names:
        used_names.add(filename)
        return filename

    if "." in filename:
        base, extension = filename.rsplit(".", 1)
        extension = "." + extension
    else:
        base, extension = filename, ""

    counter = 2
    while True:
        candidate = f"{base}_{counter}{extension}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def process_all_sheets(template_file_bytes: bytes, source_file_bytes: bytes, sheet_names: list[str]):
    """
    Behandler alle faner i kildefilen.

    Vigtigt valg:
    I stedet for at mase alle faner ind i én Excel-skabelon laver appen én færdig
    tv-appl-en-rapport pr. fane. Det er den mest sikre måde at bevare skabelonens
    struktur, layout og cellereferencer 1:1.

    Outputtet pakkes i en ZIP-fil, så brugeren kan downloade det hele på én gang.
    """
    zip_buffer = io.BytesIO()
    summaries = []
    errors = []
    used_zip_names = set()

    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for sheet_name in sheet_names:
            try:
                finished_file_bytes, summary = process_files(
                    template_file_bytes=template_file_bytes,
                    source_file_bytes=source_file_bytes,
                    selected_sheet=sheet_name,
                )

                output_filename = f"tv-appl-en_udfyldt_{safe_filename(sheet_name)}.xlsx"
                output_filename = make_unique_zip_name(output_filename, used_zip_names)

                zip_file.writestr(output_filename, finished_file_bytes)
                summary["output_filename"] = output_filename
                summaries.append(summary)

            except Exception as error:
                # Hvis én fane fejler, skal hele appen ikke dø.
                # Vi gemmer fejlen og fortsætter med næste fane.
                errors.append(
                    {
                        "sheet": sheet_name,
                        "error": str(error),
                    }
                )

    if not summaries:
        error_text = "; ".join(
            f"{item['sheet']}: {item['error']}" for item in errors[:5]
        )
        raise ValueError(
            "Ingen faner kunne behandles. "
            f"De første fejl var: {error_text}"
        )

    zip_buffer.seek(0)

    summary = {
        "mode": "all",
        "processed_count": len(summaries),
        "error_count": len(errors),
        "summaries": summaries,
        "errors": errors,
        "song_count_found": sum(item.get("song_count_found", 0) for item in summaries),
        "song_count_written": sum(item.get("song_count_written", 0) for item in summaries),
        "song_count_skipped": sum(item.get("song_count_skipped", 0) for item in summaries),
    }

    return zip_buffer.getvalue(), summary


def safe_filename(text: str) -> str:
    """
    Gør et fanenavn sikkert som filnavn.
    """
    text = safe_text(text)
    text = re.sub(r"[^A-Za-z0-9ÆØÅæøå._ -]+", "_", text)
    text = re.sub(r"\s+", "_", text).strip("_")
    return text or "rapport"


# ------------------------------------------------------------
# 11. STREAMLIT UI: UPLOAD, DROPDOWN, KNAP OG DOWNLOAD
# ------------------------------------------------------------

st.divider()

col1, col2 = st.columns(2)

with col1:
    template_file = st.file_uploader(
        "1. Upload skabelon",
        type=["xlsx"],
        help="Upload tv-appl-en.xlsx",
        key="template_uploader",
    )

with col2:
    source_file = st.file_uploader(
        "2. Upload kildefil",
        type=["xlsx"],
        help="Upload musikrapport-exporten med mange faner",
        key="source_uploader",
    )


selected_sheet = None
process_mode = None
sheet_names = []

if source_file is not None:
    try:
        source_bytes_for_sheets = source_file.getvalue()
        sheet_names = get_sheet_names(source_bytes_for_sheets)

        if sheet_names:
            st.caption(f"Fandt {len(sheet_names)} faner i kildefilen.")

            process_mode = st.radio(
                "3. Hvad vil du udtrække?",
                options=["Én valgt fane", "Alle faner"],
                horizontal=True,
            )

            if process_mode == "Én valgt fane":
                selected_sheet = st.selectbox(
                    "Vælg hvilken fane der skal udtrækkes fra",
                    options=sheet_names,
                    index=0,
                )
            else:
                st.info(
                    f"Alle {len(sheet_names)} faner bliver behandlet. "
                    "Du får én ZIP-fil med én udfyldt Excel-rapport pr. fane."
                )
                with st.expander("Se faner, der kommer med"):
                    for sheet_name in sheet_names:
                        st.write(f"- {sheet_name}")
        else:
            st.warning("Kildefilen indeholder ingen faner.")

    except InvalidFileException:
        st.error("Kildefilen ser ikke ud til at være en gyldig .xlsx-fil.")

    except Exception:
        st.error("Kunne ikke læse fanerne i kildefilen.")
        with st.expander("Vis tekniske detaljer"):
            st.code(traceback.format_exc())


st.divider()

button_disabled = (
    template_file is None
    or source_file is None
    or not sheet_names
    or process_mode is None
    or (process_mode == "Én valgt fane" and selected_sheet is None)
)

if st.button("🚀 Udfyld Rapport", type="primary", disabled=button_disabled):
    try:
        with st.spinner("Udfylder rapporten..."):
            template_bytes = template_file.getvalue()
            source_bytes = source_file.getvalue()

            if process_mode == "Alle faner":
                finished_file_bytes, summary = process_all_sheets(
                    template_file_bytes=template_bytes,
                    source_file_bytes=source_bytes,
                    sheet_names=sheet_names,
                )
                output_filename = "tv-appl-en_udfyldt_alle_faner.zip"
                download_mime = "application/zip"
                download_label = "⬇️ Download ZIP med alle udfyldte rapporter"

            else:
                finished_file_bytes, summary = process_files(
                    template_file_bytes=template_bytes,
                    source_file_bytes=source_bytes,
                    selected_sheet=selected_sheet,
                )
                summary["mode"] = "single"
                output_filename = f"tv-appl-en_udfyldt_{safe_filename(selected_sheet)}.xlsx"
                download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                download_label = "⬇️ Download udfyldt Excel-fil"

            # Gem resultatet i session_state, så download-knappen stadig virker
            # efter Streamlit har genindlæst siden.
            st.session_state["finished_file_bytes"] = finished_file_bytes
            st.session_state["output_filename"] = output_filename
            st.session_state["download_mime"] = download_mime
            st.session_state["download_label"] = download_label
            st.session_state["summary"] = summary

        st.success("Rapporten er udfyldt og klar til download.")

    except InvalidFileException:
        st.error("En af filerne ser ikke ud til at være en gyldig .xlsx-fil.")

    except Exception:
        st.error("Der skete en fejl under udfyldningen af rapporten.")
        st.write("Tjek at skabelonen og kildefilen har den forventede struktur.")

        with st.expander("Vis tekniske detaljer"):
            st.code(traceback.format_exc())


# ------------------------------------------------------------
# 12. VIS RESULTAT OG DOWNLOAD-KNAP
# ------------------------------------------------------------

if "finished_file_bytes" in st.session_state:
    summary = st.session_state.get("summary", {})

    st.subheader("✅ Klar til download")

    if summary.get("mode") == "all":
        st.write("**Udtræk:** Alle faner")
        st.write(f"**Faner behandlet:** {summary.get('processed_count', 0)}")
        st.write(f"**Faner med fejl:** {summary.get('error_count', 0)}")
        st.write(f"**Musiknumre fundet i alt:** {summary.get('song_count_found', 0)}")
        st.write(f"**Musiknumre skrevet i alt:** {summary.get('song_count_written', 0)}")

        if summary.get("song_count_skipped", 0) > 0:
            st.write(f"**Musiknumre sprunget over i alt:** {summary.get('song_count_skipped', 0)}")

        with st.expander("Se behandlede faner"):
            for item in summary.get("summaries", []):
                st.write(
                    f"- **{item.get('selected_sheet', '')}** → "
                    f"{item.get('output_filename', '')} "
                    f"({item.get('song_count_written', 0)} musiknumre skrevet)"
                )

        if summary.get("errors"):
            st.warning(
                "Nogle faner kunne ikke behandles. "
                "De fungerende faner er stadig med i ZIP-filen."
            )
            with st.expander("Se fejl pr. fane"):
                for item in summary.get("errors", []):
                    st.write(f"- **{item.get('sheet', '')}**: {item.get('error', '')}")

        all_warnings = []
        for item in summary.get("summaries", []):
            for warning in item.get("warnings", []):
                all_warnings.append(f"{item.get('selected_sheet', '')}: {warning}")

        for warning in all_warnings:
            st.warning(warning)

    else:
        st.write(f"**Valgt fane:** {summary.get('selected_sheet', '')}")
        st.write(f"**Musiknumre fundet i kildefilen:** {summary.get('song_count_found', 0)}")
        st.write(f"**Musiknumre skrevet i skabelonen:** {summary.get('song_count_written', 0)}")

        if summary.get("song_count_skipped", 0) > 0:
            st.write(f"**Musiknumre sprunget over:** {summary.get('song_count_skipped', 0)}")

        warnings = summary.get("warnings", [])
        for warning in warnings:
            st.warning(warning)

        metadata = summary.get("metadata", {})
        with st.expander("Se metadata, der blev overført"):
            st.write(metadata)

    st.download_button(
        label=st.session_state.get("download_label", "⬇️ Download fil"),
        data=st.session_state["finished_file_bytes"],
        file_name=st.session_state["output_filename"],
        mime=st.session_state.get("download_mime", "application/octet-stream"),
        type="primary",
    )


# ------------------------------------------------------------
# 13. LILLE HJÆLPETEKST TIL BRUGEREN
# ------------------------------------------------------------

with st.expander("Hvad gør appen helt præcist?"):
    st.markdown(
        """
Appen gør følgende:

1. Læser alle faner i kildefilen.
2. Lader dig vælge enten én fane eller alle faner.
3. Finder metadatafelterne:
   - Production Company
   - Production Title
   - Episode Number
   - Network/Station
4. Finder tabellen med musikdata via kolonnen **Seq #**.
5. Grupperer linjerne under hver sang, så komponister/forfattere tilhører den rigtige sang.
6. Udtrækker efternavne og fjerner foreningskoder som `[PRS]`, `[BMI]` og `(BMI)`.
7. Splitter **Music Duration** til minutter og sekunder.
8. Oversætter **Music Source** til skabelonens track types.
9. Skriver værdierne ind i de eksisterende celler i skabelonen.
10. Giver dig enten én færdig Excel-fil eller en ZIP-fil med én Excel-rapport pr. fane direkte i browseren.
"""
    )
