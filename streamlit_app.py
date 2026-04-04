"""
Werkbare Dagen Calculator — Streamlit webapp
"""

import sys
import os
from datetime import date, datetime

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "tools"))

from calculate_workdays import (
    calculate,
    STATUS_WERKBAAR, STATUS_WEEKEND, STATUS_FEESTDAG,
    STATUS_MOGELIJK_ONWERKBAAR, STATUS_ONWERKBAAR_WEER,
)

# ------------------------------------------------------------------ Config

st.set_page_config(
    page_title="Werkbare Dagen Calculator",
    page_icon="🏗️",
    layout="wide",
)

STATUS_KLEUREN = {
    STATUS_WERKBAAR:           "#d4edda",
    STATUS_WEEKEND:            "#e2e3e5",
    STATUS_FEESTDAG:           "#fff3cd",
    STATUS_MOGELIJK_ONWERKBAAR: "#ffe8b0",
    STATUS_ONWERKBAAR_WEER:    "#f8d7da",
}

STATUS_ICONEN = {
    STATUS_WERKBAAR:           "✓",
    STATUS_WEEKEND:            "—",
    STATUS_FEESTDAG:           "🎉",
    STATUS_MOGELIJK_ONWERKBAAR: "⚠️",
    STATUS_ONWERKBAAR_WEER:    "⛈️",
}


def ms_to_beaufort(ms: float) -> int:
    schaal = [0.3, 1.6, 3.4, 5.5, 8.0, 10.8, 13.9, 17.2, 20.8, 24.5, 28.5, 32.7]
    for bft, grens in enumerate(schaal):
        if ms < grens:
            return bft
    return 12


def kleur_rij(row):
    kleur = STATUS_KLEUREN.get(row["_status"], "#ffffff")
    return [f"background-color: {kleur}"] * len(row)


# ------------------------------------------------------------------ Header

st.title("🏗️ Werkbare Dagen Calculator")
st.caption("CAO Bouwnijverheid — Bft 7 wind / >20mm of >5mm/u neerslag / -3°C vorst")

# ------------------------------------------------------------------ Sidebar invoer

with st.sidebar:
    st.header("Invoer")

    startdatum = st.date_input(
        "Startdatum",
        value=date(date.today().year, 1, 1),
        format="DD-MM-YYYY",
    )
    einddatum = st.date_input(
        "Einddatum",
        value=date.today(),
        format="DD-MM-YYYY",
    )
    locatie = st.text_input("Locatie (stad)", value="Amsterdam")

    werksoort = st.radio(
        "Werksoort",
        options=["Standaard bouw", "Hijswerkzaamheden"],
        help="Hijswerk hanteert een strengere windnorm: Bft 6 (≥10,8 m/s) i.p.v. Bft 7"
    )
    norm = "hijswerk" if werksoort == "Hijswerkzaamheden" else "standaard"

    bereken = st.button("▶ Bereken", type="primary", use_container_width=True)

    st.divider()

    with st.expander("ℹ️ Definitie niet-werkbare dag"):
        wind_norm = "Beaufort 6 (≥ 10,8 m/s) — hijsnorm" if norm == "hijswerk" else "Beaufort 7 (≥ 13,9 m/s)"
        st.markdown(f"""
**Bron: Bouw & Infra CAO**

**Neerslag** — dag is niet-werkbaar als:
- Totale neerslag tijdens werkdag **> 20 mm**, en/of
- Neerslag **> 5 mm/u** gedurende **≥ 2 uur**

*Mogelijk niet-werkbaar* bij totale neerslag **10–20 mm** (beperkt uitvoerbaar)

**Wind** — ≥ {wind_norm}
minimaal 5 uur binnen de werkdag (07:00–17:00)

**Vorst** — temperatuur ≤ −3 °C
minimaal 5 uur binnen de werkdag

**Vaste niet-werkbare dagen:**
- Zaterdag en zondag
- Officiële Nederlandse feestdagen
""")

# ------------------------------------------------------------------ Berekening

if bereken:
    if not locatie.strip():
        st.error("Voer een locatie in.")
    elif einddatum < startdatum:
        st.error("Einddatum moet na de startdatum liggen.")
    else:
        with st.spinner(f"Weersdata ophalen voor {locatie}..."):
            try:
                result = calculate(startdatum, einddatum, locatie, norm=norm)
                st.session_state["result"] = result
            except ValueError as e:
                st.error(str(e))
            except ConnectionError as e:
                st.error(f"Verbindingsfout: {e}")
            except Exception as e:
                st.error(f"Onverwachte fout: {e}")

# ------------------------------------------------------------------ Resultaten

if "result" in st.session_state:
    result = st.session_state["result"]

    st.subheader(f"Resultaten — {result.locatie}")
    st.caption(f"{result.start.strftime('%d-%m-%Y')} t/m {result.eind.strftime('%d-%m-%Y')}")

    # Samenvatting metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📅 Totaal dagen", result.totaal)
    col2.metric("✅ Werkbaar", result.werkbaar)
    col3.metric("❌ Niet werkbaar", result.niet_werkbaar)
    col4.metric("⛈️ Weerverlet", result.onwerkbaar_weer)
    col5.metric("⚠️ Mogelijk niet-werkbaar", result.mogelijk_onwerkbaar)

    col6, col7, col8 = st.columns(3)
    col6.metric("📆 Weekenden", result.weekenden)
    col7.metric("🎉 Feestdagen", result.feestdagen)
    col8.metric("🚫 Niet werkbaar excl. weekend", result.feestdagen + result.onwerkbaar_weer)

    st.divider()

    # Legenda
    leg_cols = st.columns(5)
    for i, (status, kleur) in enumerate(STATUS_KLEUREN.items()):
        icoon = STATUS_ICONEN[status]
        leg_cols[i].markdown(
            f'<span style="background:{kleur};padding:2px 8px;border-radius:4px;">'
            f'{icoon} {status}</span>',
            unsafe_allow_html=True
        )

    st.markdown("")

    # Filter
    filter_col1, filter_col2 = st.columns([1, 2])
    with filter_col1:
        status_filter = st.selectbox(
            "Filter op status",
            options=["Alle", STATUS_WERKBAAR, STATUS_WEEKEND, STATUS_FEESTDAG,
                     STATUS_MOGELIJK_ONWERKBAAR, STATUS_ONWERKBAAR_WEER]
        )
    with filter_col2:
        tekst_filter = st.text_input("Zoek op datum of reden", placeholder="bijv. Maandag of Neerslag")

    # Tabel opbouwen
    fmt_wind = lambda v: f"Bft {ms_to_beaufort(v)}" if v is not None else "—"
    fmt_mm   = lambda v: f"{v} mm" if v is not None else "—"
    fmt_temp = lambda v: f"{v} °C" if v is not None else "—"

    rijen = []
    for d in result.dagen:
        rijen.append({
            "Datum":                 d.datum.strftime("%d-%m-%Y"),
            "Dag":                   d.dag_naam,
            "Status":                f"{STATUS_ICONEN.get(d.status, '')}  {d.status}",
            "Reden":                 d.reden,
            "Gem. wind":             fmt_wind(d.wind_gem_ms),
            "Totale neerslag":       fmt_mm(d.rain_totaal_mm),
            "Max neerslag (mm/u)":   fmt_mm(d.rain_max_mm),
            "Gem. temp.":            fmt_temp(d.temp_gem_c),
            "_status":               d.status,
        })

    df = pd.DataFrame(rijen)

    # Filters toepassen
    if status_filter != "Alle":
        df = df[df["_status"] == status_filter]
    if tekst_filter:
        mask = (
            df["Datum"].str.contains(tekst_filter, case=False) |
            df["Dag"].str.contains(tekst_filter, case=False) |
            df["Reden"].str.contains(tekst_filter, case=False)
        )
        df = df[mask]

    weergave = df.drop(columns=["_status"])

    styled = df.style.apply(kleur_rij, axis=1).hide(axis="index")
    # Verberg _status kolom in styled output
    styled = df.drop(columns=["_status"]).style.apply(
        lambda row: [f"background-color: {STATUS_KLEUREN.get(df.loc[row.name, '_status'], '#ffffff')}"] * len(row),
        axis=1
    ).hide(axis="index")

    st.dataframe(
        weergave,
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    st.caption(f"{len(df)} dagen weergegeven")
