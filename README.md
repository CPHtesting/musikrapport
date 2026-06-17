# Musikrapport → TV application

En lille Streamlit-app, der overfører musikrapport-data fra en Excel-export med mange faner til skabelonen `tv-appl-en.xlsx`.

## Hvad kan appen?

- Upload en skabelonfil (`tv-appl-en.xlsx`)
- Upload en kildefil med musikdata
- Vælg enten:
  - én bestemt fane, eller
  - alle faner på én gang
- Download resultatet direkte i browseren

Hvis du vælger **én valgt fane**, får du én færdig `.xlsx`-fil.

Hvis du vælger **alle faner**, får du en `.zip`-fil med én færdig Excel-rapport pr. fane.

## Kør lokalt

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy på Streamlit Community Cloud

1. Upload `app.py`, `requirements.txt` og `README.md` til et GitHub-repository.
2. Gå til Streamlit Community Cloud.
3. Vælg repository.
4. Sæt branch til `main`.
5. Sæt main file path til `app.py`.
6. Tryk Deploy.

## Vigtigt

Læg ikke rigtige Excel-rapporter eller skabeloner offentligt på GitHub, hvis de indeholder følsomme data. Appen er lavet til, at filerne uploades i browseren.
