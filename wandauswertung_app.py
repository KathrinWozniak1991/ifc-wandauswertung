import io
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element as el
import ifcopenshell.util.shape as shp
import ifcopenshell.util.unit as un
from openpyxl.styles import Font, PatternFill

# Farben: Platzhalter, bitte mit dem Vivawest-Logo abgleichen
BLAU, GRUEN, HELL = "#004B87", "#7AB51D", "#EEF3F8"

st.set_page_config(page_title="Wandauswertung", layout="wide")
st.markdown(f"""<style>
.kopf{{background:{BLAU};color:#fff;padding:18px 24px;border-radius:10px;border-bottom:5px solid {GRUEN};margin-bottom:18px}}
.kopf h1{{color:#fff;margin:0;font-size:1.6rem}}
.karte{{background:{HELL};border-left:6px solid {GRUEN};padding:14px 18px;border-radius:8px}}
.karte span{{color:#555}} .karte b{{display:block;color:{BLAU};font-size:2rem}}
[data-testid="stTable"] th{{background:{BLAU}!important;color:#fff!important}}
h3{{color:{BLAU}}}
</style>""", unsafe_allow_html=True)

logo = Path(__file__).with_name("logo.png")  # optional: Logo neben diese Datei legen
if logo.exists():
    st.image(str(logo), width=180)
st.markdown('<div class="kopf"><h1>Wandauswertung aus IFC-Modell</h1></div>', unsafe_allow_html=True)


def berechne(wand, settings, skala):
    """Fehlende Seitenfläche: Volumen / Dicke (Dicke aus Materialschichten, sonst Wandname)."""
    d = None
    try:
        sch = el.get_material(wand, should_skip_usage=True)
        if sch is not None and sch.is_a("IfcMaterialLayerSet"):
            d = sum(s.LayerThickness for s in sch.MaterialLayers) * skala
    except Exception:
        pass
    if not d:
        t = re.search(r"Wall-(\d+)", wand.Name or "")
        d = int(t.group(1)) / 1000 if t else None
    if not d:
        return None
    try:
        return shp.get_volume(ifcopenshell.geom.create_shape(settings, wand).geometry) / d
    except Exception:
        return None


@st.cache_data(show_spinner="IFC wird ausgewertet …")
def auswerten(daten: bytes, rechnen: bool):
    with tempfile.NamedTemporaryFile(suffix=".ifc", delete=False) as f:
        f.write(daten)
    m = ifcopenshell.open(f.name)
    skala, settings, zeilen = un.calculate_unit_scale(m), ifcopenshell.geom.settings(), []
    for w in m.by_type("IfcWall"):
        ps = el.get_psets(w)
        q = ps.get("Qto_WallBaseQuantities") or ps.get("BaseQuantities") or {}
        fl, herkunft = q.get("NetSideArea"), "Modell"
        if not isinstance(fl, (int, float)):
            fl, herkunft = None, "fehlt"
            if rechnen:
                fl = berechne(w, settings, skala)
                herkunft = "berechnet" if fl is not None else "fehlt"
        g = el.get_container(w)
        zeilen.append({"Name": w.Name or "", "GlobalId": w.GlobalId,
                       "Geschoss": g.Name if g else "ohne Geschoss",
                       "ist_aussen": ps.get("Pset_WallCommon", {}).get("IsExternal") is True,
                       "Fläche": fl, "Herkunft": herkunft})
    return pd.DataFrame(zeilen)


with st.sidebar:
    st.header("Einstellungen")
    datei = st.file_uploader("IFC-Modell hochladen", type="ifc")
    rechnen = st.checkbox("Fehlende Flächen aus Geometrie berechnen", value=True)

if not datei:
    st.info("Bitte links ein IFC-Modell hochladen.")
    st.stop()

df = auswerten(datei.getvalue(), rechnen)
geschosse = sorted(df["Geschoss"].unique())

with st.sidebar:
    dämm = st.multiselect("Geschosse für Fassadendämmung", geschosse,
                          default=[g for g in geschosse if re.search("erdgeschoss|obergeschoss", g.lower())])
    fund_kw = st.text_input("Fundament erkennen an (Geschossname)", "Foundation, Fundament")
    sock_kw = st.text_input("Sockel erkennen an (Wandname)", "Sockel")
    df["Sockelmaterial"] = df["Name"].str.extract(r"Sockel\.([^:]+)")[0]
    sock_mat = st.multiselect("Sockelschichten für Sockeldämmung",
                              sorted(df["Sockelmaterial"].dropna().unique()),
                              default=sorted(df["Sockelmaterial"].dropna().unique()))
    st.caption("Achtung: Sockelschichten können sich überlagern, ggf. einzelne abwählen.")


def kategorie(r):
    if not r.ist_aussen:
        return "Innenwand"
    if any(k.strip().lower() in r.Geschoss.lower() for k in fund_kw.split(",") if k.strip()):
        return "Fundament (unterirdisch)"
    if sock_kw.strip().lower() in r.Name.lower():
        return "Sockel (oberirdisch)"
    return "Außenwand Fassade" if r.Geschoss in dämm else "Außenwand sonstige"


df["Kategorie"] = df.apply(kategorie, axis=1)
df["F_ber"] = df["Fläche"].where(df["Herkunft"] == "berechnet", 0)
df["ohne"] = df["Fläche"].isna()

fassade = df.loc[df.Kategorie == "Außenwand Fassade"]
sockel = df.loc[(df.Kategorie == "Sockel (oberirdisch)") & df.Sockelmaterial.isin(sock_mat)]

c1, c2, c3 = st.columns(3)
c1.markdown(f'<div class="karte"><span>Fassadendämmung (netto)</span><b>{fassade.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="karte"><span>Sockeldämmung (netto)</span><b>{sockel.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="karte"><span>Wände im Modell</span><b>{len(df)}</b></div>', unsafe_allow_html=True)

fehlt = int(fassade.ohne.sum() + sockel.ohne.sum())
if fehlt:
    st.warning(f"{fehlt} Wände der Dämmflächen haben keine Fläche und fehlen in der Summe.")
if (df.Herkunft == "berechnet").any():
    st.caption("Teile der Flächen sind aus der Geometrie berechnet (Spalte „Herkunft“). Netto = Öffnungen abgezogen.")

zus = df.groupby("Kategorie").agg(Anzahl=("Name", "size"), Fläche=("Fläche", "sum"),
                                  berechnet=("F_ber", "sum"), ohne_Fläche=("ohne", "sum"))
zus.loc["Gesamt"] = zus.sum()
zus = zus.rename(columns={"Fläche": "Fläche (m²)", "berechnet": "davon berechnet (m²)", "ohne_Fläche": "Wände ohne Fläche"}).reset_index()

df["innen"], df["aussen"] = ~df.ist_aussen, df.ist_aussen
df["sock"], df["fund"] = df.Kategorie.str.startswith("Sockel"), df.Kategorie.str.startswith("Fundament")
df["F_innen"], df["F_aussen"] = df.Fläche.where(df.innen, 0), df.Fläche.where(df.aussen, 0)
geo = df.groupby("Geschoss").agg(Innenwände=("innen", "sum"), Außenwände=("aussen", "sum"),
                                 davon_Sockel=("sock", "sum"), davon_Fundament=("fund", "sum"),
                                 Fläche_innen=("F_innen", "sum"), Fläche_außen=("F_aussen", "sum"))
geo.loc["Gesamt"] = geo.sum()
geo = geo.rename(columns={"davon_Sockel": "davon Sockel", "davon_Fundament": "davon Fundament",
                          "Fläche_innen": "Fläche innen (m²)", "Fläche_außen": "Fläche außen (m²)"}).reset_index()

detail = df[["Name", "GlobalId", "Geschoss", "Kategorie", "Fläche", "Herkunft"]].rename(columns={"Fläche": "Fläche (m²)"})
fmt = lambda t: t.style.hide(axis="index").format(precision=1)

st.markdown("### Zusammenfassung")
st.table(fmt(zus))
st.markdown("### Pro Geschoss")
st.table(fmt(geo))
with st.expander("Alle Wände einzeln"):
    st.dataframe(detail, use_container_width=True, hide_index=True)

ergebnis = pd.DataFrame({"Ergebnis": ["Fassadendämmung netto (m²)", "Sockeldämmung netto (m²)"],
                         "Wert": [round(fassade.Fläche.sum(), 1), round(sockel.Fläche.sum(), 1)]})
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    for name, d in {"Ergebnis": ergebnis, "Zusammenfassung": zus, "Pro Geschoss": geo, "Alle Wände": detail}.items():
        d.to_excel(w, sheet_name=name, index=False)
    for ws in w.book.worksheets:
        for c in ws[1]:
            c.fill, c.font = PatternFill("solid", fgColor=BLAU[1:]), Font(color="FFFFFF", bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(60, max(len(str(c.value or "")) for c in col) + 3)
st.download_button("Als Excel exportieren", buf.getvalue(), "wandauswertung.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
