# Dagbok – Arvid

Streamlit-app for å søke og bla gjennom Arvids arbeidsdagbok (`Dagbok_2026.xlsx`).
Passordbeskyttet før dataen lastes.

> **Merk:** Selve Excel-filen ligger i et **separat privat repo**
> (`MetisPrometheus/iw-cowidagbok-data`). Dette repoet inneholder bare kode og
> kan derfor være offentlig.

## Funksjoner

- 🔍 **Søk** på tvers av alle år — *Inneholder*, *Alle ord* eller *Fuzzy* (for relaterte/feilstavede ord)
- 🎯 Filtre: år, oppdrag, datoområde, timer, oppfølging
- ✨ Treffene markeres i teksten
- 📅 Bla gjennom hvert ark som tabell
- 📊 Oversikt: innførsler og timer per år, topp 20 oppdrag
- ⬇ Eksporter treff til CSV
- 🔒 Passord-gate via `st.secrets` — dataen leses ikke før innlogging

## Lokal kjøring (uv)

Legg `Dagbok_2026.xlsx` ved siden av `app.py` (lokal fil tar forrang over GitHub-fetch):

```bash
cp /sti/til/Dagbok_2026.xlsx ./Dagbok_2026.xlsx
uv sync                                                       # lager .venv og installerer
cp .streamlit/secrets.toml.example .streamlit/secrets.toml    # sett et passord
uv run streamlit run app.py
```

> Krever [uv](https://docs.astral.sh/uv/). Avhengighetene er låst i `uv.lock`.
> Excel-filen er gitignored — den blir aldri commitet.

## Deploy til Streamlit Community Cloud

1. Push koden til GitHub (kan være offentlig — ingen data her).
2. Sørg for at `MetisPrometheus/iw-cowidagbok-data` (privat) inneholder
   `Dagbok_2026.xlsx`.
3. Lag en GitHub fine-grained PAT med **kun** lesetilgang til data-repoet:
   - Resource owner: `MetisPrometheus`
   - Repository access: kun `iw-cowidagbok-data`
   - Permissions: *Contents → Read*, *Metadata → Read*
4. Gå til https://share.streamlit.io og koble til kode-repoet.
   `Main file path`: `app.py`.
5. I appens **Secrets**-meny, lim inn:

   ```toml
   app_password = "ditt-hemmelige-passord"

   data_repo    = "MetisPrometheus/iw-cowidagbok-data"
   data_path    = "Dagbok_2026.xlsx"
   data_ref     = "main"
   github_token = "github_pat_…"   # PAT-en fra steg 3
   ```

6. Deploy. Første besøk møter login-skjermen; etter innlogging henter appen
   xlsx-filen fra det private repoet via PAT-en og cacher den i 10 minutter.

## Bytte ut Excel-filen senere

Push ny `Dagbok_2026.xlsx` til `iw-cowidagbok-data` (samme filnavn, branch
`main`). Streamlit-cachen forfaller automatisk innen 10 minutter, eller umiddelbart
hvis du restarter appen fra Streamlit Cloud-dashbordet.
