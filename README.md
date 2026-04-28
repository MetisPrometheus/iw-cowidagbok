# Dagbok – Arvid

Streamlit-app for å søke og bla gjennom Arvids arbeidsdagbok (`Dagbok_2026.xlsx`).
Passordbeskyttet før dataen lastes.

## Funksjoner

- 🔍 **Søk** på tvers av alle år — *Inneholder*, *Alle ord* eller *Fuzzy* (for relaterte/feilstavede ord)
- 🎯 Filtre: år, oppdrag, datoområde, timer, oppfølging
- ✨ Treffene markeres i teksten
- 📅 Bla gjennom hvert ark som tabell
- 📊 Oversikt: innførsler og timer per år, topp 20 oppdrag
- ⬇ Eksporter treff til CSV
- 🔒 Passord-gate via `st.secrets` — dataen leses ikke før innlogging

## Lokal kjøring (uv)

```bash
uv sync                                                       # lager .venv og installerer
cp .streamlit/secrets.toml.example .streamlit/secrets.toml    # sett et passord
uv run streamlit run app.py
```

> Krever [uv](https://docs.astral.sh/uv/). Avhengighetene er låst i `uv.lock`.

`Dagbok_2026.xlsx` må ligge ved siden av `app.py`.

## Deploy til Streamlit Community Cloud

1. Push repoet til GitHub (helst privat, siden det inneholder Excel-filen).
2. Gå til https://share.streamlit.io og koble til repoet.
3. Sett `Main file path` til `app.py`.
4. I appens **Secrets**-meny, lim inn:

   ```toml
   app_password = "ditt-hemmelige-passord"
   ```

5. Deploy. Første besøk møter login-skjermen før dataen lastes.

## Bytte ut Excel-filen senere

Last opp ny `Dagbok_2026.xlsx` til repoet (samme filnavn) og push. Streamlit Cloud
redeployer automatisk; cachen forfaller når filen endres.
