# Musikrapport → TV application

En Streamlit-app, der overfører musikrapport-data fra en Excel-export med mange faner til skabelonen `tv-appl-en.xlsx`.

## Hvad kan appen?

- Upload en skabelonfil (`tv-appl-en.xlsx`)
- Upload en kildefil med musikdata
- Vælg land/NCB-prisliste:
  - Danmark
  - Finland
  - Island
  - Norge
  - Sverige
- Vælg produktionstype ud fra det valgte land
- Vælg enten:
  - én bestemt fane, eller
  - alle faner på én gang
- Ved alle faner kan du vælge eksportform:
  - **Samlet musikrapport**: én Excel-fil, hvor alle tracks fra alle faner samles i den samme `Music content`-tabel
  - **Én musikrapport pr. fane**: én ZIP-fil med separate Excel-rapporter
- Automatisk samling af gentagne bumpers/vignetter, fx norske bumpers, så de skrives som `Bumpernavn x antal`
- Download resultatet direkte i browseren

## Smart-regler

Appen har en simpel regelmotor, der bruger det valgte land og den valgte produktionstype.

Eksempel: Hvis du vælger **Norge** og appen finder gentagne bumpers, samler den dem automatisk til én linje, lægger varigheden sammen og skriver titlen med `x antal`.

Eksempel:

```text
THE VOICE OF…. Bumpers x 8
```

Det gør rapporten hurtigere at implementere i det system, hvor den bagefter skal uploades.

## Kør lokalt

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy på Streamlit Community Cloud

1. Upload `app.py`, `requirements.txt`, `README.md` og `.gitignore` til et GitHub-repository.
2. Gå til Streamlit Community Cloud.
3. Vælg repository.
4. Sæt branch til `main`.
5. Sæt main file path til `app.py`.
6. Tryk Deploy.

## Vigtigt

Læg ikke rigtige Excel-rapporter eller skabeloner offentligt på GitHub, hvis de indeholder følsomme data. Appen er lavet til, at filerne uploades i browseren.
