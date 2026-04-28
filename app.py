"""Dagbok dashboard for Arvid — search across all yearly diary sheets."""
from __future__ import annotations

import hmac
import html
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

APP_TITLE = "Dagbok – Arvid"
XLSX_NAME = "Dagbok_2026.xlsx"
CANONICAL_COLS = ["Dato", "Oppdrag", "Uke", "Dagbok", "Følges_opp", "Timer"]


# ──────────────────────────────────────────────────────────────────────────
# Page config + styling
# ──────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📓",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
html, body, [class*="css"]  { font-size: 17px; }
.block-container { padding-top: 2rem; max-width: 1300px; }
h1, h2, h3 { letter-spacing: -0.01em; }
mark { background: #fff3a3; padding: 0 2px; border-radius: 3px; }
.entry-card {
    background: #f7f9fc;
    border: 1px solid #e3e8ef;
    border-radius: 12px;
    padding: 16px 18px;
    margin-bottom: 12px;
}
.entry-meta {
    color: #5b6472;
    font-size: 0.9rem;
    margin-bottom: 6px;
    display: flex;
    gap: 14px;
    flex-wrap: wrap;
}
.entry-meta b { color: #1c1f26; }
.entry-body { font-size: 1.02rem; line-height: 1.55; color: #1c1f26; }
.followup-pill {
    display: inline-block;
    background: #ffe7c2;
    color: #8a4b00;
    border-radius: 999px;
    padding: 1px 9px;
    font-size: 0.78rem;
    font-weight: 600;
}
.login-card {
    max-width: 420px;
    margin: 8vh auto 0 auto;
    padding: 32px 28px 28px 28px;
    border-radius: 16px;
    background: #f7f9fc;
    border: 1px solid #e3e8ef;
    text-align: center;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────
# Auth gate
# ──────────────────────────────────────────────────────────────────────────
def _check_password() -> bool:
    if st.session_state.get("authed"):
        return True

    expected = st.secrets.get("app_password")
    if not expected:
        st.error(
            "Mangler passord-konfigurasjon. Sett `app_password` i "
            "`.streamlit/secrets.toml` (lokalt) eller i Streamlit Cloud sin Secrets-meny."
        )
        st.stop()

    st.markdown('<div class="login-card">', unsafe_allow_html=True)
    st.markdown(f"### 📓 {APP_TITLE}")
    st.caption("Skriv inn passord for å åpne dagboken.")

    with st.form("login", clear_on_submit=False):
        pw = st.text_input("Passord", type="password", label_visibility="collapsed",
                           placeholder="Passord")
        submitted = st.form_submit_button("Logg inn", use_container_width=True)

    if submitted:
        if hmac.compare_digest(pw or "", str(expected)):
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Feil passord.")

    st.markdown("</div>", unsafe_allow_html=True)
    return False


# ──────────────────────────────────────────────────────────────────────────
# Data loading
# ──────────────────────────────────────────────────────────────────────────
def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {}
    for col in df.columns:
        c = str(col).strip()
        cl = c.lower()
        if cl == "dato":
            rename[col] = "Dato"
        elif cl == "oppdrag":
            rename[col] = "Oppdrag"
        elif cl in ("uke nr.", "uke nr", "uke"):
            rename[col] = "Uke"
        elif cl == "dagbok":
            rename[col] = "Dagbok"
        elif cl == "timer":
            rename[col] = "Timer"
        elif ("lges" in cl and "opp" in cl) or "ølges" in cl:
            rename[col] = "Følges_opp"
    df = df.rename(columns=rename)

    cols = list(df.columns)
    # If Dagbok missing but col 3 is anonymous, treat it as the description.
    if "Dagbok" not in df.columns and len(cols) > 3 and str(cols[3]).startswith("Unnamed"):
        df = df.rename(columns={cols[3]: "Dagbok"})
        cols = list(df.columns)
    # If Timer missing but col 5 is anonymous, treat it as hours.
    if "Timer" not in df.columns and len(cols) > 5 and str(cols[5]).startswith("Unnamed"):
        df = df.rename(columns={cols[5]: "Timer"})

    keep = [c for c in CANONICAL_COLS if c in df.columns]
    return df[keep]


def _extract_year(sheet_name: str):
    m = re.search(r"(\d{4})", sheet_name)
    return int(m.group(1)) if m else sheet_name


def _clean_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    if df.empty:
        return df
    # Drop trailing all-NaN columns (bloated 2004 sheet has hundreds of them).
    df = df.dropna(axis=1, how="all")
    df = _normalize_columns(df)
    if df.empty:
        return df
    # Drop rows missing both date and description.
    has_dato = df["Dato"].notna() if "Dato" in df.columns else pd.Series(False, index=df.index)
    has_dagbok = df["Dagbok"].notna() if "Dagbok" in df.columns else pd.Series(False, index=df.index)
    df = df[has_dato | has_dagbok].copy()
    if df.empty:
        return df
    # Coerce types.
    if "Dato" in df.columns:
        df["Dato"] = pd.to_datetime(df["Dato"], errors="coerce")
    if "Uke" in df.columns:
        df["Uke"] = pd.to_numeric(df["Uke"], errors="coerce").astype("Int64")
    if "Timer" in df.columns:
        df["Timer"] = pd.to_numeric(df["Timer"], errors="coerce")
    if "Følges_opp" in df.columns:
        df["Følges_opp"] = df["Følges_opp"].notna() & (df["Følges_opp"].astype(str).str.strip() != "")
    if "Dagbok" in df.columns:
        df["Dagbok"] = df["Dagbok"].astype("string").fillna("")
    if "Oppdrag" in df.columns:
        df["Oppdrag"] = df["Oppdrag"].astype("string").fillna("")
    df["Ark"] = sheet_name
    df["År"] = _extract_year(sheet_name)
    return df.reset_index(drop=True)


@st.cache_data(show_spinner="Laster dagbok…")
def load_data(xlsx_path: str, mtime: float):
    """Load every sheet, clean, return (master_df, per_sheet_dict)."""
    xl = pd.ExcelFile(xlsx_path)
    per_sheet: dict[str, pd.DataFrame] = {}
    cleaned: list[pd.DataFrame] = []
    for name in xl.sheet_names:
        try:
            raw = pd.read_excel(xlsx_path, sheet_name=name)
        except Exception:
            continue
        clean = _clean_sheet(raw, name)
        if clean.empty:
            continue
        per_sheet[name] = clean
        cleaned.append(clean)
    if not cleaned:
        return pd.DataFrame(columns=CANONICAL_COLS + ["Ark", "År"]), {}
    master = pd.concat(cleaned, ignore_index=True)
    # Stable order: oldest first
    master = master.sort_values(by=["Dato"], na_position="last").reset_index(drop=True)
    return master, per_sheet


# ──────────────────────────────────────────────────────────────────────────
# Search helpers
# ──────────────────────────────────────────────────────────────────────────
def _highlight_terms(text: str, terms: list[str]) -> str:
    safe = html.escape(text)
    for term in sorted({t for t in terms if t}, key=len, reverse=True):
        if not term:
            continue
        pattern = re.compile(re.escape(term), re.IGNORECASE)
        safe = pattern.sub(lambda m: f"<mark>{m.group(0)}</mark>", safe)
    return safe


def _highlight_fuzzy_span(text: str, query: str) -> str:
    if not text or not query:
        return html.escape(text or "")
    try:
        align = fuzz.partial_ratio_alignment(query.lower(), text.lower())
    except Exception:
        return html.escape(text)
    if align is None or align.score < 60:
        return html.escape(text)
    start, end = align.dest_start, align.dest_end
    return (
        html.escape(text[:start])
        + "<mark>"
        + html.escape(text[start:end])
        + "</mark>"
        + html.escape(text[end:])
    )


def filter_search(df: pd.DataFrame, query: str, mode: str, fuzzy_threshold: int) -> pd.DataFrame:
    if not query.strip():
        return df
    text = df["Dagbok"].fillna("").astype(str)
    oppdrag = df["Oppdrag"].fillna("").astype(str)
    haystack = (text + " ⏐ " + oppdrag).str.lower()
    q = query.strip().lower()

    if mode == "Inneholder":
        mask = haystack.str.contains(re.escape(q), regex=True, na=False)
    elif mode == "Alle ord":
        tokens = [t for t in q.split() if t]
        mask = pd.Series(True, index=df.index)
        for t in tokens:
            mask &= haystack.str.contains(re.escape(t), regex=True, na=False)
    else:  # Fuzzy
        scores = haystack.map(lambda h: fuzz.partial_ratio(q, h) if h else 0)
        mask = scores >= fuzzy_threshold

    return df[mask].copy()


# ──────────────────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────────────────
def _format_date(d) -> str:
    if pd.isna(d):
        return "—"
    return pd.Timestamp(d).strftime("%Y-%m-%d")


def _format_year(y) -> str:
    return str(y) if not isinstance(y, float) else str(int(y))


def render_search_tab(master: pd.DataFrame):
    st.markdown("### 🔍 Søk i alle år")
    col_q, col_mode, col_thr = st.columns([5, 2, 2])
    with col_q:
        query = st.text_input(
            "Søk", placeholder="Søk etter ord, prosjektnummer, navn …",
            label_visibility="collapsed",
        )
    with col_mode:
        mode = st.selectbox("Modus", ["Inneholder", "Alle ord", "Fuzzy"],
                            label_visibility="collapsed")
    with col_thr:
        threshold = st.slider("Likhet", 50, 100, 75, disabled=(mode != "Fuzzy"),
                              label_visibility="collapsed")

    # Sidebar filters
    with st.sidebar:
        st.markdown("### Filter")
        years = sorted({y for y in master["År"].unique() if y is not None}, key=lambda x: str(x))
        sel_years = st.multiselect("År", years, default=years)

        oppdrag_options = sorted(master["Oppdrag"].dropna().unique().tolist())
        sel_oppdrag = st.multiselect("Oppdrag", oppdrag_options, default=[])

        date_min = master["Dato"].min()
        date_max = master["Dato"].max()
        date_range = None
        if pd.notna(date_min) and pd.notna(date_max):
            date_range = st.date_input(
                "Datoområde",
                value=(date_min.date(), date_max.date()),
                min_value=date_min.date(),
                max_value=date_max.date(),
            )

        if "Timer" in master.columns and master["Timer"].notna().any():
            tmin, tmax = float(master["Timer"].min()), float(master["Timer"].max())
            if tmin == tmax:
                tmax = tmin + 1
            timer_range = st.slider("Timer", min_value=float(tmin), max_value=float(tmax),
                                    value=(float(tmin), float(tmax)))
        else:
            timer_range = None

        only_followup = st.checkbox("Kun rader merket for oppfølging")

    df = master.copy()
    if sel_years:
        df = df[df["År"].isin(sel_years)]
    if sel_oppdrag:
        df = df[df["Oppdrag"].isin(sel_oppdrag)]
    if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
        d0, d1 = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
        df = df[(df["Dato"].isna()) | ((df["Dato"] >= d0) & (df["Dato"] < d1))]
    if timer_range is not None and "Timer" in df.columns:
        lo, hi = timer_range
        df = df[(df["Timer"].isna()) | ((df["Timer"] >= lo) & (df["Timer"] <= hi))]
    if only_followup and "Følges_opp" in df.columns:
        df = df[df["Følges_opp"] == True]

    df = filter_search(df, query, mode, threshold)

    # Result header
    left, right = st.columns([3, 1])
    with left:
        st.caption(f"**{len(df):,}** treff".replace(",", " "))
    with right:
        if not df.empty:
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇ Last ned CSV", csv,
                               file_name="dagbok_treff.csv", mime="text/csv",
                               use_container_width=True)

    if df.empty:
        st.info("Ingen treff. Prøv et kortere søk eller bytt til *Fuzzy*-modus.")
        return

    # Render result cards (cap to 300 for performance / readability)
    MAX_CARDS = 300
    shown = df.head(MAX_CARDS)
    if len(df) > MAX_CARDS:
        st.warning(f"Viser de første {MAX_CARDS} treffene. Bruk filter for å snevre inn.")

    terms = [t for t in re.split(r"\s+", query.strip()) if t] if mode != "Fuzzy" else []

    for _, row in shown.iterrows():
        dato = _format_date(row.get("Dato"))
        år = _format_year(row.get("År", ""))
        uke = row.get("Uke")
        uke_str = f"Uke {int(uke)}" if pd.notna(uke) else ""
        oppdrag = row.get("Oppdrag", "") or ""
        timer = row.get("Timer")
        timer_str = f"{timer:g} t" if pd.notna(timer) else ""
        followup = bool(row.get("Følges_opp", False))
        text = str(row.get("Dagbok", "") or "")
        if mode == "Fuzzy" and query.strip():
            body = _highlight_fuzzy_span(text, query)
        else:
            body = _highlight_terms(text, terms)
        meta_bits = [
            f"<b>{dato}</b>",
            f"År {år}" if år else "",
            uke_str,
            f"Oppdrag <b>{html.escape(oppdrag)}</b>" if oppdrag else "",
            timer_str,
        ]
        meta_bits = [b for b in meta_bits if b]
        if followup:
            meta_bits.append('<span class="followup-pill">Følg opp</span>')
        st.markdown(
            f'<div class="entry-card">'
            f'<div class="entry-meta">{" · ".join(meta_bits)}</div>'
            f'<div class="entry-body">{body}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )


def render_browse_tab(master: pd.DataFrame, per_sheet: dict[str, pd.DataFrame]):
    st.markdown("### 📅 Bla gjennom år")
    sheets = list(per_sheet.keys())
    if not sheets:
        st.info("Ingen ark å vise.")
        return
    # Default to most recent dagbok sheet
    default_idx = next((i for i, s in enumerate(sheets) if "2026" in s), len(sheets) - 1)
    sheet = st.selectbox("Velg ark", sheets, index=default_idx)
    df = per_sheet[sheet].copy()

    # Quick stats
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Antall rader", f"{len(df):,}".replace(",", " "))
    with c2:
        total_t = df["Timer"].sum() if "Timer" in df.columns else 0
        st.metric("Totalt timer", f"{total_t:,.1f}".replace(",", " "))
    with c3:
        if "Følges_opp" in df.columns:
            st.metric("Til oppfølging", int(df["Følges_opp"].sum()))
        else:
            st.metric("Til oppfølging", 0)

    if "Oppdrag" in df.columns and "Timer" in df.columns:
        top = (df.groupby("Oppdrag")["Timer"].sum()
               .sort_values(ascending=False).head(5).reset_index())
        if not top.empty:
            with st.expander("Topp 5 oppdrag (timer)", expanded=False):
                st.dataframe(top, use_container_width=True, hide_index=True)

    show_cols = [c for c in ["Dato", "Uke", "Oppdrag", "Dagbok", "Timer", "Følges_opp"] if c in df.columns]
    st.dataframe(
        df[show_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Dato": st.column_config.DateColumn("Dato", format="YYYY-MM-DD"),
            "Dagbok": st.column_config.TextColumn("Dagbok", width="large"),
            "Timer": st.column_config.NumberColumn("Timer", format="%.2f"),
            "Følges_opp": st.column_config.CheckboxColumn("Følg opp"),
        },
        height=600,
    )


def render_overview_tab(master: pd.DataFrame):
    st.markdown("### 📊 Oversikt")
    if master.empty:
        st.info("Ingen data.")
        return

    by_year_count = (master.groupby("År").size()
                     .rename("Antall").reset_index().sort_values("År", key=lambda s: s.astype(str)))
    by_year_hours = (master.groupby("År")["Timer"].sum().fillna(0)
                     .rename("Timer").reset_index().sort_values("År", key=lambda s: s.astype(str)))

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Innførsler per år**")
        st.bar_chart(by_year_count, x="År", y="Antall", height=320)
    with c2:
        st.markdown("**Timer per år**")
        st.bar_chart(by_year_hours, x="År", y="Timer", height=320)

    st.markdown("**Topp 20 oppdrag (totale timer på tvers av år)**")
    top = (master.groupby("Oppdrag")["Timer"].sum()
           .sort_values(ascending=False).head(20).reset_index())
    st.dataframe(top, use_container_width=True, hide_index=True,
                 column_config={"Timer": st.column_config.NumberColumn("Timer", format="%.1f")})


# ──────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────
def main():
    if not _check_password():
        return

    xlsx_path = Path(__file__).parent / XLSX_NAME
    if not xlsx_path.exists():
        st.error(f"Fant ikke `{XLSX_NAME}` ved siden av appen. Last opp filen til "
                 "repoet og deploy på nytt.")
        st.stop()

    master, per_sheet = load_data(str(xlsx_path), xlsx_path.stat().st_mtime)

    with st.sidebar:
        st.markdown(f"### 📓 {APP_TITLE}")
        st.caption(
            f"{len(master):,} innførsler · {len(per_sheet)} ark"
            .replace(",", " ")
        )
        if st.button("Logg ut", use_container_width=True):
            st.session_state.pop("authed", None)
            st.rerun()
        st.divider()

    tab_search, tab_browse, tab_overview = st.tabs(
        ["🔍 Søk", "📅 Bla gjennom år", "📊 Oversikt"]
    )
    with tab_search:
        render_search_tab(master)
    with tab_browse:
        render_browse_tab(master, per_sheet)
    with tab_overview:
        render_overview_tab(master)


if __name__ == "__main__":
    main()
