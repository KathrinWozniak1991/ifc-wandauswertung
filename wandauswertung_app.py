import io
import re
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element as el
import ifcopenshell.util.shape as shp
import ifcopenshell.util.unit as un
from openpyxl.styles import Font, PatternFill

# Farben: VIVAWEST-Design (Dunkelblau, Grün, Hellgrau)
BLAU, GRUEN, HELL = "#004B87", "#7AB51D", "#EEF3F8"

st.set_page_config(page_title="bauteilauswertung-ifc", layout="wide")
st.markdown(f"""<style>
.kopf{{background:{BLAU};color:#fff;padding:22px 26px;border-radius:10px;border-bottom:5px solid {GRUEN};margin-bottom:20px}}
.kopf h1{{color:#fff;margin:0;font-size:1.75rem}}
.kopf p{{color:#E2ECF5;margin:6px 0 0 0;font-size:1rem}}
.karte{{background:{HELL};border-left:6px solid {GRUEN};padding:14px 18px;border-radius:8px;margin-bottom:10px}}
.karte span{{color:#555;font-size:0.9rem}} .karte b{{display:block;color:{BLAU};font-size:1.8rem}}
.info-box{{background:#FFFFFF;border:1px solid #D8E2EC;border-top:4px solid {BLAU};padding:18px 20px;border-radius:8px;height:100%;box-shadow:0 2px 4px rgba(0,0,0,0.04)}}
.info-box h4{{color:{BLAU};margin-top:0;margin-bottom:8px;font-size:1.1rem}}
.info-box p{{color:#334E68;font-size:0.95rem;line-height:1.5;margin:0}}
.ueber-mich{{background:{HELL};border-left:6px solid {BLAU};padding:20px 24px;border-radius:8px;margin-top:20px}}
.ueber-mich h3{{margin-top:0;color:{BLAU}}}
.ueber-mich p{{color:#243B53;font-size:0.96rem;line-height:1.6;margin-bottom:8px}}
[data-testid="stTable"] th{{background:{BLAU}!important;color:#fff!important}}
h3{{color:{BLAU}}}
</style>""", unsafe_allow_html=True)

logo = Path(__file__).with_name("logo.png")
if logo.exists():
    st.image(str(logo), width=180)
st.markdown(
    '<div class="kopf">'
    '<h1>bauteilauswertung-ifc</h1>'
    '<p>Automatisierte Mengen- und Flächenermittlung aus OpenBIM-Bestandsmodellen</p>'
    '</div>',
    unsafe_allow_html=True
)


def berechne_wand(wand, settings, skala):
    """Fehlende Seitenfläche der Wand über Geometrie-Volumen und Dicke berechnen."""
    d = None
    try:
        # 1. Versuch: Gesamtdicke aus den Materialschichten (IfcMaterialLayerSet) auslesen
        sch = el.get_material(wand, should_skip_usage=True)
        if sch is not None and sch.is_a("IfcMaterialLayerSet"):
            d = sum(s.LayerThickness for s in sch.MaterialLayers) * skala
    except Exception:
        pass
    if not d:
        # 2. Versuch: Dicke aus dem Wandnamen extrahieren (z. B. "Wall-240" -> 0,24 m)
        t = re.search(r"Wall-(\d+)", wand.Name or "")
        d = int(t.group(1)) / 1000 if t else None
    if not d:
        return None
    try:
        # Seitenfläche = Volumen / Wanddicke
        return shp.get_volume(ifcopenshell.geom.create_shape(settings, wand).geometry) / d
    except Exception:
        return None


def berechne_dach_geometrie(bauteil, settings):
    """Fallback: Dachfläche (Oberseite) und Neigung aus der 3D-Geometrie ermitteln."""
    try:
        geom = ifcopenshell.geom.create_shape(settings, bauteil).geometry
        verts = np.array(geom.verts).reshape(-1, 3)
        faces = np.array(geom.faces).reshape(-1, 3)
        flaeche_oben = 0.0
        neigungen = []

        for f in faces:
            p0, p1, p2 = verts[f[0]], verts[f[1]], verts[f[2]]
            kreuz = np.cross(p1 - p0, p2 - p0)
            norm_len = np.linalg.norm(kreuz)
            if norm_len == 0:
                continue
            normal = kreuz / norm_len
            # Nur Dreiecke berücksichtigen, deren Normale nach oben zeigt (Z > 0.1)
            if normal[2] > 0.1:
                tri_flaeche = 0.5 * norm_len
                flaeche_oben += tri_flaeche
                winkel = np.degrees(np.arccos(np.clip(normal[2], -1.0, 1.0)))
                neigungen.append((winkel, tri_flaeche))

        if flaeche_oben <= 0:
            return None, None
        # Flächengewichtete mittlere Dachneigung berechnen
        mittlere_neigung = sum(w * a for w, a in neigungen) / flaeche_oben
        return flaeche_oben, mittlere_neigung
    except Exception:
        return None, None


@st.cache_data(show_spinner="IFC wird ausgewertet … (Wände, Fenster & Dächer)")
def auswerten(daten: bytes, rechnen: bool):
    with tempfile.NamedTemporaryFile(suffix=".ifc", delete=False) as f:
        f.write(daten)
    m = ifcopenshell.open(f.name)
    skala = un.calculate_unit_scale(m)
    settings = ifcopenshell.geom.settings() if rechnen else None

    # 1. Wände auswerten (IfcWall)
    zeilen_waende = []
    for w in m.by_type("IfcWall"):
        ps = el.get_psets(w)
        q = ps.get("Qto_WallBaseQuantities") or ps.get("BaseQuantities") or {}
        fl, herkunft = q.get("NetSideArea"), "Modell"
        if not isinstance(fl, (int, float)):
            fl, herkunft = None, "fehlt"
            if rechnen:
                fl = berechne_wand(w, settings, skala)
                herkunft = "berechnet" if fl is not None else "fehlt"
        g = el.get_container(w)
        zeilen_waende.append({
            "Name": w.Name or "",
            "GlobalId": w.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "ist_aussen": ps.get("Pset_WallCommon", {}).get("IsExternal") is True,
            "Fläche": fl,
            "Herkunft": herkunft
        })

    # 2. Fenster auswerten (IfcWindow)
    zeilen_fenster = []
    for f in m.by_type("IfcWindow"):
        ps = el.get_psets(f)
        b = getattr(f, "OverallWidth", None)
        h = getattr(f, "OverallHeight", None)
        b = b * skala if b is not None else None
        h = h * skala if h is not None else None

        if b is None or h is None:
            q = ps.get("Qto_WindowBaseQuantities") or ps.get("BaseQuantities") or {}
            b = b or q.get("Width")
            h = h or q.get("Height")

        fl = (b * h) if (b and h) else None
        g = el.get_container(f)

        zeilen_fenster.append({
            "Name": f.Name or "Fenster",
            "GlobalId": f.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "ist_aussen": ps.get("Pset_WindowCommon", {}).get("IsExternal") is True,
            "Breite (m)": b,
            "Höhe (m)": h,
            "Fläche (m²)": fl
        })

    # 3. Dächer auswerten (IfcRoof und IfcSlab mit PredefinedType == ROOF)
    zeilen_daecher = []
    dach_kandidaten = []

    # Alle IfcRoof prüfen: Wenn ein IfcRoof Unterelemente (Teile) hat, nehmen wir nur die Teile (Doppelzählung vermeiden)
    for r in m.by_type("IfcRoof"):
        teile = el.get_parts(r)
        if teile:
            for t in teile:
                if t not in dach_kandidaten:
                    dach_kandidaten.append(t)
        else:
            if r not in dach_kandidaten:
                dach_kandidaten.append(r)

    # Zusätzlich alle Deckenplatten (IfcSlab) aufnehmen, die als Dach (ROOF) deklariert sind
    for s in m.by_type("IfcSlab"):
        if getattr(s, "PredefinedType", None) == "ROOF" and s not in dach_kandidaten:
            dach_kandidaten.append(s)

    for d in dach_kandidaten:
        ps = el.get_psets(d)
        q = (
            ps.get("Qto_RoofBaseQuantities")
            or ps.get("Qto_SlabBaseQuantities")
            or ps.get("BaseQuantities")
            or {}
        )
        # Nettofläche bevorzugen, sonst Bruttofläche
        fl = q.get("NetArea") if isinstance(q.get("NetArea"), (int, float)) else q.get("GrossArea")
        herkunft = "Modell" if isinstance(fl, (int, float)) else "fehlt"

        # Dachneigung aus Pset_RoofCommon oder Pset_SlabCommon auslesen
        pset_common = ps.get("Pset_RoofCommon") or ps.get("Pset_SlabCommon") or {}
        neigung = pset_common.get("PitchAngle")

        # Fallback über 3D-Geometrie, falls Fläche oder Neigung im Modell fehlt
        if (fl is None or neigung is None) and rechnen:
            fl_geo, neig_geo = berechne_dach_geometrie(d, settings)
            if fl is None and fl_geo is not None:
                fl = fl_geo
                herkunft = "berechnet"
            if neigung is None and neig_geo is not None:
                neigung = neig_geo

        # Dachform bestimmen (Flachdach < 5° Neigung vs. Geneigtes Dach)
        if isinstance(neigung, (int, float)):
            dachform = "Flachdach" if neigung < 5.0 else "Geneigtes Dach"
        else:
            dachform = "Unbekannt"

        g = el.get_container(d)
        zeilen_daecher.append({
            "Name": d.Name or "Dachfläche",
            "IFC-Klasse": d.is_a(),
            "GlobalId": d.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "Dachform": dachform,
            "Neigung (°)": round(neigung, 1) if isinstance(neigung, (int, float)) else None,
            "Fläche (m²)": fl,
            "Herkunft": herkunft
        })

    return pd.DataFrame(zeilen_waende), pd.DataFrame(zeilen_fenster), pd.DataFrame(zeilen_daecher)


with st.sidebar:
    st.header("Einstellungen")
    datei = st.file_uploader("IFC-Modell hochladen", type="ifc")
    rechnen = st.checkbox("Fehlende Flächen aus Geometrie berechnen", value=True)

# --- STARTSEITE (wird angezeigt, solange noch keine IFC-Datei hochgeladen wurde) ---
if not datei:
    st.markdown("### Willkommen bei bauteilauswertung-ifc")
    st.write(
        "Diese Web-Anwendung wertet digitale Gebäudemodelle im offenen **IFC-Format (OpenBIM)** "
        "automatisiert aus. Sie wurde speziell entwickelt, um Bestandsgebäude im Wohnungsbau "
        "schnell, transparent und zuverlässig für **Modernisierungs- und Sanierungsmaßnahmen** "
        "(z. B. Fassadendämmung, Fenstertausch und Dachsanierung) zu analysieren."
    )

    # 3 Vorteilskarten nebeneinander
    v1, v2, v3 = st.columns(3)
    with v1:
        st.markdown(
            '<div class="info-box">'
            '<h4>1. Schnelle Mengen statt Handarbeit</h4>'
            '<p>Kein mühsames manuelles Zusammensuchen in CAD- oder Viewer-Programmen: '
            'Fassaden-, Sockel-, Fenster- und Dachflächen stehen wenige Sekunden nach dem Upload '
            'übersichtlich nach Geschossen sortiert bereit.</p>'
            '</div>',
            unsafe_allow_html=True
        )
    with v2:
        st.markdown(
            '<div class="info-box">'
            '<h4>2. Intelligente Geometrie-Prüfung</h4>'
            '<p>Fehlen in einem Bestandsmodell einmal Flächenangaben in den Eigenschaften, '
            'berechnet die App diese automatisch direkt aus der 3D-Geometrie nach – und weist '
            'transparent aus, welche Werte aus dem Modell stammen.</p>'
            '</div>',
            unsafe_allow_html=True
        )
    with v3:
        st.markdown(
            '<div class="info-box">'
            '<h4>3. Fertiger Excel-Bericht</h4>'
            '<p>Mit einem Klick lassen sich alle Zusammenfassungen und Einzelnachweise als '
            'sauber formatierte Excel-Datei herunterladen – direkt nutzbar für '
            'Kostenschätzungen, Ausschreibungen und Projektunterlagen.</p>'
            '</div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### So einfach funktioniert's")
    st.markdown(
        "1. **IFC-Datei auswählen:** Klicken Sie links in der Seitenleiste auf **Browse files** und wählen Sie ein `.ifc`-Modell aus.\n"
        "2. **Filter anpassen:** Legen Sie bei Bedarf fest, welche Geschosse zur Fassadendämmung zählen oder wie Sockel und Fundamente erkannt werden.\n"
        "3. **Ergebnisse prüfen & exportieren:** Prüfen Sie die Kennzahlen direkt im Browser und laden Sie unten den fertigen **Excel-Bericht** herunter."
    )

    # Bereich "Über mich"
    st.markdown(
        '<div class="ueber-mich">'
        '<h3>Über mich & das Projekt</h3>'
        '<p><strong>Entwickelt von Kathrin Wozniak</strong> – Architektin mit Fokus auf Bestandsmodernisierung und OpenBIM im Wohnungsbau.</p>'
        '<p>In der täglichen Planungspraxis liegen Bestandsgebäude heute oft als IFC-Modelle vor, '
        'doch für schnelle Wirtschaftlichkeits- und Sanierungsbetrachtungen fehlen häufig passgenaue, '
        'unkomplizierte Auswertungswerkzeuge ohne komplexe Spezialsoftware. '
        'Aus diesem Bedarf heraus – und im Rahmen meiner Weiterbildung im Bereich <strong>BIM (Building Information Modeling)</strong> – '
        'ist <em>bauteilauswertung-ifc</em> entstanden: ein praxisnahes Werkzeug aus der Architektur für die Architektur, '
        'das offene Gebäudedaten direkt im Browser nutzbar macht.</p>'
        '</div>',
        unsafe_allow_html=True
    )
    st.stop()
# --- ENDE STARTSEITE ---

df_waende, df_fenster, df_daecher = auswerten(datei.getvalue(), rechnen)
geschosse_waende = sorted(df_waende["Geschoss"].unique())

with st.sidebar:
    st.subheader("Wandauswertung")
    dämm = st.multiselect(
        "Geschosse für Fassadendämmung",
        geschosse_waende,
        default=[g for g in geschosse_waende if re.search("erdgeschoss|obergeschoss", g.lower())]
    )
    fund_kw = st.text_input("Fundament erkennen an (Geschossname)", "Foundation, Fundament")
    sock_kw = st.text_input("Sockel erkennen an (Wandname)", "Sockel")
    df_waende["Sockelmaterial"] = df_waende["Name"].str.extract(r"Sockel\.([^:]+)")[0]
    sock_mat = st.multiselect(
        "Sockelschichten für Sockeldämmung",
        sorted(df_waende["Sockelmaterial"].dropna().unique()),
        default=sorted(df_waende["Sockelmaterial"].dropna().unique())
    )
    st.caption("Achtung: Sockelschichten können sich überlagern, ggf. einzelne abwählen.")


def kategorie(r):
    if not r.ist_aussen:
        return "Innenwand"
    if any(k.strip().lower() in r.Geschoss.lower() for k in fund_kw.split(",") if k.strip()):
        return "Fundament (unterirdisch)"
    if sock_kw.strip().lower() in r.Name.lower():
        return "Sockel (oberirdisch)"
    return "Außenwand Fassade" if r.Geschoss in dämm else "Außenwand sonstige"


df_waende["Kategorie"] = df_waende.apply(kategorie, axis=1)
df_waende["F_ber"] = df_waende["Fläche"].where(df_waende["Herkunft"] == "berechnet", 0)
df_waende["ohne"] = df_waende["Fläche"].isna()

fassade = df_waende.loc[df_waende.Kategorie == "Außenwand Fassade"]
sockel = df_waende.loc[(df_waende.Kategorie == "Sockel (oberirdisch)") & df_waende.Sockelmaterial.isin(sock_mat)]

dach_gesamt = df_daecher["Fläche (m²)"].sum() if not df_daecher.empty else 0.0
fenster_gesamt = df_fenster["Fläche (m²)"].sum() if not df_fenster.empty else 0.0

# Kennzahlen-Karten in 4 Spalten
c1, c2, c3, c4 = st.columns(4)
c1.markdown(f'<div class="karte"><span>Fassadendämmung (netto)</span><b>{fassade.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="karte"><span>Sockeldämmung (netto)</span><b>{sockel.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="karte"><span>Fensterfläche ({len(df_fenster)} Stk.)</span><b>{fenster_gesamt:,.1f} m²</b></div>', unsafe_allow_html=True)
c4.markdown(f'<div class="karte"><span>Dachfläche ({len(df_daecher)} Teilfl.)</span><b>{dach_gesamt:,.1f} m²</b></div>', unsafe_allow_html=True)

fehlt = int(fassade.ohne.sum() + sockel.ohne.sum())
if fehlt:
    st.warning(f"{fehlt} Wände der Dämmflächen haben keine Fläche und fehlen in der Summe.")
if (df_waende.Herkunft == "berechnet").any() or (not df_daecher.empty and (df_daecher.Herkunft == "berechnet").any()):
    st.caption("Teile der Flächen wurden aus der 3D-Geometrie berechnet (Spalte „Herkunft“).")

# Zusammenfassung Wände
zus_waende = df_waende.groupby("Kategorie").agg(
    Anzahl=("Name", "size"),
    Fläche=("Fläche", "sum"),
    berechnet=("F_ber", "sum"),
    ohne_Fläche=("ohne", "sum")
)
zus_waende.loc["Gesamt"] = zus_waende.sum()
zus_waende = zus_waende.rename(columns={
    "Fläche": "Fläche (m²)",
    "berechnet": "davon berechnet (m²)",
    "ohne_Fläche": "Wände ohne Fläche"
}).reset_index()

# Zusammenfassung Fenster pro Geschoss
if not df_fenster.empty:
    zus_fenster = df_fenster.groupby("Geschoss").agg(
        Anzahl=("Name", "size"),
        Fläche=("Fläche (m²)", "sum")
    ).reset_index()
    zus_fenster.loc[len(zus_fenster)] = ["Gesamt", zus_fenster["Anzahl"].sum(), zus_fenster["Fläche"].sum()]
else:
    zus_fenster = pd.DataFrame(columns=["Geschoss", "Anzahl", "Fläche (m²)"])

# Zusammenfassung Dächer nach Dachform
if not df_daecher.empty:
    zus_daecher = df_daecher.groupby("Dachform").agg(
        Anzahl=("Name", "size"),
        Fläche=("Fläche (m²)", "sum")
    ).reset_index()
    zus_daecher.loc[len(zus_daecher)] = ["Gesamt", zus_daecher["Anzahl"].sum(), zus_daecher["Fläche"].sum()]
else:
    zus_daecher = pd.DataFrame(columns=["Dachform", "Anzahl", "Fläche (m²)"])

df_waende["innen"], df_waende["aussen"] = ~df_waende.ist_aussen, df_waende.ist_aussen
df_waende["sock"], df_waende["fund"] = df_waende.Kategorie.str.startswith("Sockel"), df_waende.Kategorie.str.startswith("Fundament")
df_waende["F_innen"], df_waende["F_aussen"] = df_waende.Fläche.where(df_waende.innen, 0), df_waende.Fläche.where(df_waende.aussen, 0)
geo_waende = df_waende.groupby("Geschoss").agg(
    Innenwände=("innen", "sum"),
    Außenwände=("aussen", "sum"),
    davon_Sockel=("sock", "sum"),
    davon_Fundament=("fund", "sum"),
    Fläche_innen=("F_innen", "sum"),
    Fläche_außen=("F_aussen", "sum")
)
geo_waende.loc["Gesamt"] = geo_waende.sum()
geo_waende = geo_waende.rename(columns={
    "davon_Sockel": "davon Sockel",
    "davon_Fundament": "davon Fundament",
    "Fläche_innen": "Fläche innen (m²)",
    "Fläche_außen": "Fläche außen (m²)"
}).reset_index()

detail_waende = df_waende[["Name", "GlobalId", "Geschoss", "Kategorie", "Fläche", "Herkunft"]].rename(columns={"Fläche": "Fläche (m²)"})
detail_fenster = df_fenster[["Name", "GlobalId", "Geschoss", "ist_aussen", "Breite (m)", "Höhe (m)", "Fläche (m²)"]] if not df_fenster.empty else pd.DataFrame()
detail_daecher = df_daecher[["Name", "IFC-Klasse", "GlobalId", "Geschoss", "Dachform", "Neigung (°)", "Fläche (m²)", "Herkunft"]] if not df_daecher.empty else pd.DataFrame()

fmt = lambda t: t.style.hide(axis="index").format(precision=1)

st.markdown("### Wände: Zusammenfassung")
st.table(fmt(zus_waende))
st.markdown("### Wände: Pro Geschoss")
st.table(fmt(geo_waende))

st.markdown("### Fenster: Zusammenfassung pro Geschoss")
if not zus_fenster.empty:
    st.table(fmt(zus_fenster))
else:
    st.info("Keine Fenster (IfcWindow) im Modell gefunden.")

st.markdown("### Dächer: Zusammenfassung")
if not zus_daecher.empty:
    st.table(fmt(zus_daecher))
else:
    st.info("Keine Dachflächen (IfcRoof / IfcSlab mit PredefinedType ROOF) im Modell gefunden.")

with st.expander("Alle Wände einzeln"):
    st.dataframe(detail_waende, use_container_width=True, hide_index=True)
with st.expander("Alle Fenster einzeln"):
    st.dataframe(detail_fenster, use_container_width=True, hide_index=True)
with st.expander("Alle Dächer einzeln"):
    st.dataframe(detail_daecher, use_container_width=True, hide_index=True)

ergebnis = pd.DataFrame({
    "Ergebnis": [
        "Fassadendämmung netto (m²)",
        "Sockeldämmung netto (m²)",
        "Fensterfläche gesamt (m²)",
        "Dachfläche gesamt (m²)"
    ],
    "Wert": [
        round(fassade.Fläche.sum(), 1),
        round(sockel.Fläche.sum(), 1),
        round(fenster_gesamt, 1),
        round(dach_gesamt, 1)
    ]
})
buf = io.BytesIO()

# Strukturierter Excel-Export mit eigenen Reitern im VIVAWEST-Design
sheets = {
    "Ergebnis": ergebnis,
    "Wände Zusammenfassung": zus_waende,
    "Wände pro Geschoss": geo_waende,
    "Fenster Zusammenfassung": zus_fenster,
    "Dächer Zusammenfassung": zus_daecher,
    "Alle Wände": detail_waende,
    "Alle Fenster": detail_fenster,
    "Alle Dächer": detail_daecher
}

with pd.ExcelWriter(buf, engine="openpyxl") as w:
    for name, d in sheets.items():
        d.to_excel(w, sheet_name=name, index=False)
    for ws in w.book.worksheets:
        for c in ws[1]:
            c.fill, c.font = PatternFill("solid", fgColor=BLAU[1:]), Font(color="FFFFFF", bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(60, max(len(str(c.value or "")) for c in col) + 3)

st.download_button(
    "Als Excel exportieren",
    buf.getvalue(),
    "bauteilauswertung-ifc.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
