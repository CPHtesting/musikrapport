# Musikrapport → TV application

En simpel Streamlit-app, der overfører data fra en musikrapport-export med mange faner til NCB-skabelonen `tv-appl-en.xlsx`.

## Hvad appen gør

- Upload skabelonen `tv-appl-en.xlsx`
- Upload kildefilen med musikdata
- Vælg fane fra kildefilen
- Tryk **Udfyld Rapport**
- Download den udfyldte Excel-fil

Appen bruger `openpyxl` og forsøger at bevare skabelonens layout, struktur og formatering ved kun at skrive værdier ind i eksisterende celler.

## Filer i projektet

```text
app.py
requirements.txt
README.md
.gitignore
```

## Kør appen lokalt

Installer først pakkerne:

```bash
pip install -r requirements.txt
```

Start appen:

```bash
streamlit run app.py
```

## Deploy på Streamlit Community Cloud

1. Upload disse filer til et GitHub-repository.
2. Gå til Streamlit Community Cloud.
3. Vælg dit repository.
4. Sæt **Main file path** til:

```text
app.py
```

5. Tryk **Deploy**.

## Vigtigt

Excel-filer er ignoreret i `.gitignore`, så du ikke ved en fejl uploader skabeloner eller musikrapporter til GitHub.
Filerne skal uploades direkte i appen via browseren.
