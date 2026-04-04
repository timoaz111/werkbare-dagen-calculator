"""
Werkbare Dagen Calculator — Streamlit webapp
"""

import sys
import os
from datetime import date

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
    STATUS_WERKBAAR:            "#d4edda",
    STATUS_WEEKEND:             "#e2e3e5",
    STATUS_FEESTDAG:            "#fff3cd",
    STATUS_MOGELIJK_ONWERKBAAR: "#ffe8b0",
    STATUS_ONWERKBAAR_WEER:     "#f8d7da",
}

STATUS_ICONEN = {
    STATUS_WERKBAAR:            "✓",
    STATUS_WEEKEND:             "—",
    STATUS_FEESTDAG:            "🎉",
    STATUS_MOGELIJK_ONWERKBAAR: "⚠️",
    STATUS_ONWERKBAAR_WEER:     "⛈️",
}

# CSS: donkere header + tabelstyling
st.markdown("""
<style>
.header-box {
    background-color: #343a40;
    padding: 18px 24px 14px 24px;
    border-radius: 8px;
    margin-bottom: 18px;
}
.header-box h1 {
    color: #ffffff;
    margin: 0;
    font-size: 1.7rem;
}
.header-box p {
    color: #adb5bd;
    margin: 4px 0 0 0;
    font-size: 0.85rem;
}
.summary-box {
    background-color: #f8f9fa;
    border: 1px solid #dee2e6;
    border-radius: 8px;
    padding: 16px 20px;
    margin-bottom: 12px;
}
.summary-label {
    font-weight: bold;
    font-size: 0.82rem;
    color: #343a40;
}
.summary-value {
    font-size: 0.82rem;
    color: #343a40;
}
.legenda-item {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 4px;
    font-size: 0.8rem;
    margin-right: 8px;
    border: 1px solid #dee2e6;
}
div[data-testid="stDataFrame"] table td {
    font-size: 0.82rem !important;
}
</style>
""", unsafe_allow_html=True)


def ms_to_beaufort(ms: float) -> int:
    schaal = [0.3, 1.6, 3.4, 5.5, 8.0, 10.8, 13.9, 17.2, 20.8, 24.5, 28.5, 32.7]
    for bft, grens in enumerate(schaal):
        if ms < grens:
            return bft
    return 12


# ------------------------------------------------------------------ Header

st.markdown("""
<div class="header-box">
  <h1>🏗️ Werkbare Dagen Calculator</h1>
  <p>CAO Bouwnijverheid &mdash; Bft 7 wind / &gt;20mm of &gt;5mm/u neerslag / -3&deg;C vorst</p>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------ Invoer (horizontaal, zoals desktop)

with st.container():
    c1, c2, c3, c4, c5 = st.columns([2, 2, 2, 3, 2])
    with c1:
        startdatum = st.date_input("Startdatum", value=date(date.today().year, 1, 1), format="DD-MM-YYYY")
    with c2:
        einddatum = st.date_input("Einddatum", value=date.today(), format="DD-MM-YYYY")
    with c3:
        locatie = st.text_input("Locatie (stad)", value="Amsterdam")
    with c4:
        werksoort = st.radio(
            "Werksoort",
            options=["Standaard bouw  (wind ≥ Bft 7)", "Hijswerkzaamheden  (wind ≥ Bft 6)"],
            horizontal=True,
        )
    with c5:
        st.markdown("<br>", unsafe_allow_html=True)
        bereken = st.button("▶ Bereken", type="primary", use_container_width=True)

norm = "hijswerk" if "Hijswerk" in werksoort else "standaard"

# Info expander onder invoer
with st.expander("ℹ️ Definitie niet-werkbare dag (Bouw & Infra CAO)"):
    wind_norm = "Beaufort 6 (≥ 10,8 m/s) — hijsnorm" if norm == "hijswerk" else "Beaufort 7 (≥ 13,9 m/s)"
    st.markdown(f"""
**NEERSLAG** — dag is niet-werkbaar als:
- Totale neerslag tijdens werkdag **> 20 mm**, en/of
- Neerslag **> 5 mm/u** gedurende **≥ 2 uur**

*Mogelijk niet-werkbaar* bij totale neerslag **10–20 mm** — werkzaamheden beperkt uitvoerbaar.

**WIND** — ≥ {wind_norm}, minimaal 5 uur binnen de werkdag (07:00–17:00)

**VORST** — temperatuur ≤ −3 °C, minimaal 5 uur binnen de werkdag

**VASTE NIET-WERKBARE DAGEN:** zaterdag, zondag en officiële Nederlandse feestdagen
(Nieuwjaar, Pasen, Koningsdag, Bevrijdingsdag, Hemelvaartsdag, Pinksteren, Kerst)
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

    st.markdown(f"### Samenvatting — {result.locatie}")
    st.caption(f"{result.start.strftime('%d-%m-%Y')} t/m {result.eind.strftime('%d-%m-%Y')}")

    # Samenvatting tabel in stijl van het origineel
    def samenvatting_rij(label, waarde, vet=False):
        stijl = "font-weight:bold;" if vet else ""
        return f"""
        <tr>
          <td style="padding:3px 16px 3px 8px;{stijl} font-size:0.83rem;">{label}</td>
          <td style="padding:3px 8px;font-size:0.83rem;">{waarde}</td>
        </tr>"""

    col_s1, col_s2, col_s3 = st.columns(3)

    with col_s1:
        st.markdown(f"""
<div class="summary-box">
<table>
{samenvatting_rij("Totaal kalenderdagen:", result.totaal, vet=True)}
{samenvatting_rij("Werkbare dagen:", result.werkbaar, vet=True)}
{samenvatting_rij("Niet werkbare dagen:", result.niet_werkbaar, vet=True)}
</table>
</div>""", unsafe_allow_html=True)

    with col_s2:
        st.markdown(f"""
<div class="summary-box">
<table>
{samenvatting_rij("↳ Weekenden:", result.weekenden)}
{samenvatting_rij("↳ Feestdagen:", result.feestdagen)}
{samenvatting_rij("↳ Onwerkbaar weer:", result.onwerkbaar_weer)}
{samenvatting_rij("↳ Mogelijk niet-werkbaar:", result.mogelijk_onwerkbaar)}
</table>
</div>""", unsafe_allow_html=True)

    with col_s3:
        st.markdown(f"""
<div class="summary-box">
<table>
{samenvatting_rij("Niet werkbaar (excl. weekend):", result.feestdagen + result.onwerkbaar_weer, vet=True)}
{samenvatting_rij("Weerverlet dagen:", result.onwerkbaar_weer, vet=True)}
</table>
</div>""", unsafe_allow_html=True)

    # Legenda
    legenda_html = "".join([
        f'<span class="legenda-item" style="background-color:{kleur};">'
        f'{STATUS_ICONEN[status]} {status}</span>'
        for status, kleur in STATUS_KLEUREN.items()
    ])
    st.markdown(legenda_html, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # ------------------------------------------------------------------ Filter + tabel

    st.markdown("#### Details per dag")
    fc1, fc2 = st.columns([1, 2])
    with fc1:
        status_filter = st.selectbox(
            "Filter op status",
            options=["Alle", STATUS_WERKBAAR, STATUS_WEEKEND, STATUS_FEESTDAG,
                     STATUS_MOGELIJK_ONWERKBAAR, STATUS_ONWERKBAAR_WEER],
            label_visibility="collapsed",
        )
    with fc2:
        tekst_filter = st.text_input(
            "Zoeken",
            placeholder="Zoek op datum, dag of reden...",
            label_visibility="collapsed",
        )

    fmt_wind = lambda v: f"Bft {ms_to_beaufort(v)}" if v is not None else "—"
    fmt_mm   = lambda v: f"{v} mm" if v is not None else "—"
    fmt_temp = lambda v: f"{v} °C" if v is not None else "—"

    rijen = []
    for d in result.dagen:
        rijen.append({
            "Datum":               d.datum.strftime("%d-%m-%Y"),
            "Dag":                 d.dag_naam,
            "Status":              f"{STATUS_ICONEN.get(d.status, '')}  {d.status}",
            "Reden / Opmerking":   d.reden,
            "Gem. wind":           fmt_wind(d.wind_gem_ms),
            "Totale neerslag":     fmt_mm(d.rain_totaal_mm),
            "Max neerslag (mm/u)": fmt_mm(d.rain_max_mm),
            "Gem. temp.":          fmt_temp(d.temp_gem_c),
            "_status":             d.status,
        })

    df = pd.DataFrame(rijen)

    if status_filter != "Alle":
        df = df[df["_status"] == status_filter]
    if tekst_filter:
        mask = (
            df["Datum"].str.contains(tekst_filter, case=False) |
            df["Dag"].str.contains(tekst_filter, case=False) |
            df["Reden / Opmerking"].str.contains(tekst_filter, case=False)
        )
        df = df[mask]

    # Kleur per rij via styler
    status_kolom = df["_status"].tolist()
    weergave = df.drop(columns=["_status"]).reset_index(drop=True)

    def kleur_rijen(df_in):
        kleuren = []
        for i in range(len(df_in)):
            kleur = STATUS_KLEUREN.get(status_kolom[i], "#ffffff")
            kleuren.append([f"background-color: {kleur}"] * len(df_in.columns))
        return pd.DataFrame(kleuren, columns=df_in.columns, index=df_in.index)

    styled = weergave.style.apply(kleur_rijen, axis=None)

    st.dataframe(styled, use_container_width=True, hide_index=True, height=520)
    st.caption(f"{len(df)} dagen weergegeven")
