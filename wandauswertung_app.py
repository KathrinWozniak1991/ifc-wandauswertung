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

# ==============================================================================
# 1. DESIGN-KONFIGURATION & SEITENLAYOUT
# ==============================================================================
# Harmonische Beige- und Sandtöne (Taupe, Warmes Beige, Creme, Dunkelbraun)
HAUPT_BEIGE = "#8C735B"   # Warmes Taupe-Braun für Kopfzeile & Tabellenköpfe
AKZENT_SAND = "#C8B299"   # Sandbeige für Rahmen & Akzentlinien
HELL_CREME  = "#F5EFE6"   # Helles Creme-Beige für Karten & Hintergrundboxen
TEXT_BRAUN  = "#3E3228"   # Dunkles Espresso-Braun für hohen Kontrast

st.set_page_config(page_title="bauteilauswertung-ifc", layout="wide")

st.markdown(f"""<style>
.kopf{{background:{HAUPT_BEIGE};color:#fff;padding:22px 26px;border-radius:10px;border-bottom:5px solid {AKZENT_SAND};margin-bottom:20px}}
.kopf h1{{color:#fff;margin:0;font-size:1.75rem}}
.kopf p{{color:#F5EFE6;margin:6px 0 0 0;font-size:1rem}}
.karte{{background:{HELL_CREME};border-left:6px solid {HAUPT_BEIGE};padding:14px 18px;border-radius:8px;margin-bottom:10px}}
.karte span{{color:#6E5D4F;font-size:0.85rem}} .karte b{{display:block;color:{TEXT_BRAUN};font-size:1.55rem}}
.info-box{{background:#FAF7F2;border:1px solid #E5DCCB;border-top:4px solid {HAUPT_BEIGE};padding:18px 20px;border-radius:8px;height:100%;box-shadow:0 2px 4px rgba(62,50,40,0.05)}}
.info-box h4{{color:{TEXT_BRAUN};margin-top:0;margin-bottom:8px;font-size:1.1rem}}
.info-box p{{color:#5C4D40;font-size:0.95rem;line-height:1.5;margin:0}}
.ueber-mich-platzhalter{{width:145px;height:145px;border-radius:50%;border:4px solid {AKZENT_SAND};background:{HAUPT_BEIGE};color:#fff;display:flex;align-items:center;justify-content:center;font-size:2rem;font-weight:bold;flex-shrink:0}}
[data-testid="stTable"] th{{background:{HAUPT_BEIGE}!important;color:#fff!important}}
h3{{color:{TEXT_BRAUN}}}
</style>""", unsafe_allow_html=True)

st.markdown(
    '<div class="kopf">'
    '<h1>bauteilauswertung-ifc</h1>'
    '<p>Automatisierte Mengen-, Flächen- und Längenermittlung aus OpenBIM-Bestandsmodellen</p>'
    '</div>',
    unsafe_allow_html=True
)


# ==============================================================================
# 2. HILFSFUNKTIONEN (3D-GEOMETRIE-FALLBACKS)
# ==============================================================================
def berechne_wand(wand, settings, skala: float):
    """Berechnet die fehlende Seitenfläche einer Wand über Volumen / Wanddicke."""
    d = None
    try:
        # 1. Versuch: Gesamtdicke aus den Materialschichten (IfcMaterialLayerSet) auslesen
        sch = el.get_material(wand, should_skip_usage=True)
        if sch is not None and sch.is_a("IfcMaterialLayerSet"):
            d = sum(s.LayerThickness for s in sch.MaterialLayers) * skala
    except Exception:
        pass

    if not d:
        # 2. Versuch: Dicke aus dem Wandnamen extrahieren (z. B. "Wall-420" -> 0,42 m)
        t = re.search(r"Wall-(\d+)", wand.Name or "")
        d = int(t.group(1)) / 1000 if t else None

    if not d:
        return None

    try:
        # Seitenfläche = Volumen (in m³) / Wanddicke (in m)
        volumen = shp.get_volume(ifcopenshell.geom.create_shape(settings, wand).geometry)
        return volumen / d
    except Exception:
        return None


def berechne_dach_geometrie(bauteil, settings):
    """Fallback: Ermittelt Dachfläche (Oberseite) und mittlere Neigung aus der 3D-Geometrie."""
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

        mittlere_neigung = sum(w * a for w, a in neigungen) / flaeche_oben
        return flaeche_oben, mittlere_neigung
    except Exception:
        return None, None


def berechne_laenge_3d(bauteil, settings, modus: str = "horizontal"):
    """
    Fallback für Geländer, Regenrinnen und Fallrohre:
    Schätzt die Bauteillänge aus der 3D-Ausdehnung (Bounding Box) in Metern.
    """
    try:
        geom = ifcopenshell.geom.create_shape(settings, bauteil).geometry
        verts = np.array(geom.verts).reshape(-1, 3)
        if len(verts) == 0:
            return None
        ausdehnung = verts.max(axis=0) - verts.min(axis=0)
        dx, dy, dz = float(ausdehnung[0]), float(ausdehnung[1]), float(ausdehnung[2])

        if modus == "vertikal":
            return dz if dz > 0.05 else None
        if modus == "diagonal":
            horiz = max(dx, dy)
            return float(np.hypot(horiz, dz)) if horiz > 0.05 else None
        horiz = max(dx, dy)
        return horiz if horiz > 0.05 else None
    except Exception:
        return None


def hole_laenge_aus_psets(ps: dict, skala: float):
    """Sucht in allen PropertySets nach typischen Längen-Eigenschaften."""
    for _, props in ps.items():
        if not isinstance(props, dict):
            continue
        for key in ["Length", "NetLength", "GrossLength", "NominalLength", "Länge"]:
            val = props.get(key)
            if isinstance(val, (int, float)) and val > 0:
                return val * skala if val > 100 else float(val)
    return None


# ==============================================================================
# 3. HAUPTAUSWERTUNG DER IFC-DATEI
# ==============================================================================
@st.cache_data(show_spinner="IFC wird ausgewertet … (Wände, Fenster, Türen, Dächer, Geländer & Entwässerung)")
def auswerten(daten: bytes, rechnen: bool):
    with tempfile.NamedTemporaryFile(suffix=".ifc", delete=False) as f:
        f.write(daten)
        tmp_pfad = Path(f.name)

    try:
        m = ifcopenshell.open(str(tmp_pfad))
    finally:
        tmp_pfad.unlink(missing_ok=True)

    skala = un.calculate_unit_scale(m)
    settings = ifcopenshell.geom.settings() if rechnen else None

    # --- 1. Wände auswerten (IfcWall) ---
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

        wand_name = w.Name or ""
        ist_aussen = ps.get("Pset_WallCommon", {}).get("IsExternal") is True
        # Fallback für Revit-Exporte ohne IsExternal: Typische Außenwand-Begriffe prüfen
        if not ist_aussen and any(k in wand_name.lower() for k in ["fassade", "aussen", "außen", "exterior", "sockel"]):
            ist_aussen = True

        g = el.get_container(w)
        zeilen_waende.append({
            "Name": wand_name,
            "GlobalId": w.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "ist_aussen": ist_aussen,
            "Fläche": fl,
            "Herkunft": herkunft
        })

    # --- 2. Fenster auswerten (IfcWindow) ---
    zeilen_fenster = []
    for f in m.by_type("IfcWindow"):
        ps = el.get_psets(f)
        q = ps.get("Qto_WindowBaseQuantities") or ps.get("BaseQuantities") or {}

        b_raw = getattr(f, "OverallWidth", None)
        h_raw = getattr(f, "OverallHeight", None)
        b = b_raw * skala if isinstance(b_raw, (int, float)) else None
        h = h_raw * skala if isinstance(h_raw, (int, float)) else None

        if b is None and isinstance(q.get("Width"), (int, float)):
            b = q["Width"] * skala if q["Width"] > 50 else q["Width"]
        if h is None and isinstance(q.get("Height"), (int, float)):
            h = q["Height"] * skala if q["Height"] > 50 else q["Height"]

        fl = q.get("Area") if isinstance(q.get("Area"), (int, float)) else ((b * h) if (b and h) else None)
        g = el.get_container(f)

        zeilen_fenster.append({
            "Name": f.Name or "Fenster",
            "GlobalId": f.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "ist_aussen": ps.get("Pset_WindowCommon", {}).get("IsExternal") is True,
            "Breite (m)": round(b, 3) if isinstance(b, (int, float)) else None,
            "Höhe (m)": round(h, 3) if isinstance(h, (int, float)) else None,
            "Fläche (m²)": fl
        })

    # --- 3. Türen auswerten (IfcDoor) ---
    zeilen_tueren = []
    for t in m.by_type("IfcDoor"):
        ps = el.get_psets(t)
        q = ps.get("Qto_DoorBaseQuantities") or ps.get("BaseQuantities") or {}

        b_raw = getattr(t, "OverallWidth", None)
        h_raw = getattr(t, "OverallHeight", None)
        b = b_raw * skala if isinstance(b_raw, (int, float)) else None
        h = h_raw * skala if isinstance(h_raw, (int, float)) else None

        if b is None and isinstance(q.get("Width"), (int, float)):
            b = q["Width"] * skala if q["Width"] > 50 else q["Width"]
        if h is None and isinstance(q.get("Height"), (int, float)):
            h = q["Height"] * skala if q["Height"] > 50 else q["Height"]

        fl = q.get("Area") if isinstance(q.get("Area"), (int, float)) else ((b * h) if (b and h) else None)
        g = el.get_container(t)
        typ_obj = el.get_type(t)
        typ_name = typ_obj.Name if typ_obj and typ_obj.Name else ""

        # Außenlage prüfen (direktes Property + Fallback über Mutterwand)
        ist_aussen = ps.get("Pset_DoorCommon", {}).get("IsExternal") is True
        wand_name = ""
        try:
            for rel in getattr(t, "FillsVoids", []):
                wand = rel.RelatingBuildingElement
                if wand:
                    wand_name = wand.Name or ""
                    wand_ps = el.get_psets(wand)
                    if wand_ps.get("Pset_WallCommon", {}).get("IsExternal") is True:
                        ist_aussen = True
                    if any(k in wand_name.lower() for k in ["fassade", "aussen", "außen", "exterior", "sockel"]):
                        ist_aussen = True
        except Exception:
            pass

        zeilen_tueren.append({
            "Name": t.Name or "Tür",
            "Typ": typ_name,
            "Mutterwand": wand_name,
            "GlobalId": t.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "ist_aussen": ist_aussen,
            "Breite (m)": round(b, 3) if isinstance(b, (int, float)) else None,
            "Höhe (m)": round(h, 3) if isinstance(h, (int, float)) else None,
            "Fläche (m²)": fl
        })

    # --- 4. Dächer auswerten (IfcRoof und IfcSlab mit PredefinedType == ROOF) ---
    zeilen_daecher = []
    dach_kandidaten = []

    for r in m.by_type("IfcRoof"):
        teile = el.get_parts(r)
        if teile:
            for t in teile:
                if t not in dach_kandidaten:
                    dach_kandidaten.append(t)
        else:
            if r not in dach_kandidaten:
                dach_kandidaten.append(r)

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
        fl = q.get("NetArea") if isinstance(q.get("NetArea"), (int, float)) else q.get("GrossArea")
        herkunft = "Modell" if isinstance(fl, (int, float)) else "fehlt"

        pset_common = ps.get("Pset_RoofCommon") or ps.get("Pset_SlabCommon") or {}
        neigung = pset_common.get("PitchAngle")

        if (fl is None or neigung is None) and rechnen:
            fl_geo, neig_geo = berechne_dach_geometrie(d, settings)
            if fl is None and fl_geo is not None:
                fl = fl_geo
                herkunft = "berechnet"
            if neigung is None and neig_geo is not None:
                neigung = neig_geo

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

    # --- 5. Geländer auswerten (IfcRailing) ---
    zeilen_gelaender = []
    for r in m.by_type("IfcRailing"):
        ps = el.get_psets(r)
        predef = (getattr(r, "PredefinedType", None) or "").upper()
        name_str = r.Name or "Geländer"
        suchtext = f"{name_str} {predef}".lower()

        if predef == "STAIR" or any(k in suchtext for k in ["treppe", "stair", "handlauf"]):
            gelaender_art = "Treppengeländer"
            geo_modus = "diagonal"
        elif any(k in suchtext for k in ["fenster", "absturz", "franz", "brüstung", "window", "balustrade"]):
            gelaender_art = "Absturzgeländer (Fenster)"
            geo_modus = "horizontal"
        else:
            ist_aussen = ps.get("Pset_RailingCommon", {}).get("IsExternal") is True
            gelaender_art = "Absturzgeländer (Fenster)" if ist_aussen else "Treppengeländer"
            geo_modus = "horizontal" if ist_aussen else "diagonal"

        laenge = hole_laenge_aus_psets(ps, skala)
        herkunft = "Modell" if laenge is not None else "fehlt"

        if laenge is None and rechnen:
            laenge = berechne_laenge_3d(r, settings, modus=geo_modus)
            herkunft = "berechnet" if laenge is not None else "fehlt"

        g = el.get_container(r)
        zeilen_gelaender.append({
            "Name": name_str,
            "GlobalId": r.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "Geländerart": gelaender_art,
            "Länge (m)": round(laenge, 2) if isinstance(laenge, (int, float)) else None,
            "Herkunft": herkunft
        })

    # --- 6. Regenrinnen & Fallrohre auswerten (Schema-sicher für IFC2x3 & IFC4) ---
    zeilen_entwaesserung = []
    rohr_kandidaten = []
    for klasse in ["IfcFlowSegment", "IfcBuildingElementProxy"]:
        try:
            rohr_kandidaten.extend(m.by_type(klasse))
        except Exception:
            pass

    for p in rohr_kandidaten:
        ps = el.get_psets(p)
        predef = (getattr(p, "PredefinedType", None) or "").upper()
        name_str = p.Name or ""
        typ_obj = el.get_type(p)
        typ_name = typ_obj.Name if typ_obj and typ_obj.Name else ""
        suchtext = f"{name_str} {typ_name} {predef}".lower()

        ist_rinne = predef == "GUTTER" or any(k in suchtext for k in ["rinne", "dachrinne", "gutter", "kastenrinne", "halbrund"])
        ist_fallrohr = predef == "DOWNPIPE" or any(k in suchtext for k in ["fallrohr", "regenrohr", "downpipe", "rainwater"])

        if not (ist_rinne or ist_fallrohr):
            continue

        bauteil_art = "Regenrinne" if ist_rinne else "Fallrohr"
        geo_modus = "horizontal" if ist_rinne else "vertikal"

        laenge = hole_laenge_aus_psets(ps, skala)
        herkunft = "Modell" if laenge is not None else "fehlt"

        if laenge is None and rechnen:
            laenge = berechne_laenge_3d(p, settings, modus=geo_modus)
            herkunft = "berechnet" if laenge is not None else "fehlt"

        g = el.get_container(p)
        zeilen_entwaesserung.append({
            "Name": name_str or bauteil_art,
            "Typ": typ_name,
            "GlobalId": p.GlobalId,
            "Geschoss": g.Name if g else "ohne Geschoss",
            "Bauteilart": bauteil_art,
            "Länge (m)": round(laenge, 2) if isinstance(laenge, (int, float)) else None,
            "Herkunft": herkunft
        })

    return (
        pd.DataFrame(zeilen_waende),
        pd.DataFrame(zeilen_fenster),
        pd.DataFrame(zeilen_tueren),
        pd.DataFrame(zeilen_daecher),
        pd.DataFrame(zeilen_gelaender),
        pd.DataFrame(zeilen_entwaesserung)
    )


# ==============================================================================
# 4. SEITENLEISTE & STARTSEITE
# ==============================================================================
with st.sidebar:
    st.header("Einstellungen")
    datei = st.file_uploader("IFC-Modell hochladen", type="ifc")
    rechnen = st.checkbox("Fehlende Maße aus 3D-Geometrie berechnen", value=True)

if not datei:
    st.markdown("### Willkommen bei bauteilauswertung-ifc")
    st.write(
        "Diese Web-Anwendung wertet digitale Gebäudemodelle im offenen **IFC-Format (OpenBIM)** "
        "automatisiert aus. Sie wurde speziell entwickelt, um Bestandsgebäude im Wohnungsbau "
        "schnell, transparent und zuverlässig für **Modernisierungs- und Sanierungsmaßnahmen** "
        "(Fassade, Fenster, Türen, Dach, Geländer und Dachentwässerung) zu analysieren."
    )

    v1, v2, v3 = st.columns(3)
    with v1:
        st.markdown(
            '<div class="info-box">'
            '<h4>1. Schnelle Mengen statt Handarbeit</h4>'
            '<p>Kein mühsames Zusammensuchen in CAD- oder Viewer-Programmen: '
            'Fassaden-, Sockel-, Fenster-, Tür- und Dachmengen sowie Geländer- und Rinnenlängen '
            'stehen wenige Sekunden nach dem Upload bereit.</p>'
            '</div>',
            unsafe_allow_html=True
        )
    with v2:
        st.markdown(
            '<div class="info-box">'
            '<h4>2. Intelligente Geometrie-Prüfung</h4>'
            '<p>Fehlen in einem Bestandsmodell Flächen- oder Längenangaben in den Eigenschaften, '
            'berechnet die App diese automatisch aus der 3D-Geometrie nach – und weist '
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
        "2. **Filter anpassen:** Prüfen Sie in der Seitenleiste die Zuordnung von Geschossen, Sockeln und Türarten.\n"
        "3. **Ergebnisse prüfen & exportieren:** Prüfen Sie die Kennzahlen direkt im Browser und laden Sie unten den fertigen **Excel-Bericht** herunter."
    )

    # --- Bereich "Über mich & das Projekt" (gestochen scharf durch hohe Render-Auflösung) ---
    st.markdown("---")
    basis_ordner = Path(__file__).resolve().parent
    profil_pfad = next(
        (p for p in basis_ordner.iterdir() if p.is_file() and "profil" in p.stem.lower() and p.suffix.lower() in {".jpg", ".jpeg", ".png"}),
        None
    )

    # CSS: Bild wird mit 450px scharf geladen, aber im Browser exakt als 145x145px Kreis dargestellt
    st.markdown(f"""<style>
    [data-testid="stImage"] {{
        display: flex;
        justify-content: center;
    }}
    [data-testid="stImage"] img {{
        width: 145px !important;
        height: 145px !important;
        max-width: 145px !important;
        border-radius: 50%;
        border: 4px solid {AKZENT_SAND};
        object-fit: cover;
        object-position: center 20%;
        image-rendering: -webkit-optimize-contrast;
    }}
    </style>""", unsafe_allow_html=True)

    with st.container(border=True):
        col_bild, col_text = st.columns([1, 4], vertical_alignment="center")
        with col_bild:
            bild_angezeigt = False
            if profil_pfad:
                try:
                    # width=450 sorgt für 3-fache Pixeldichte (Retina-scharf statt 140px-Pixelbrei)
                    st.image(str(profil_pfad), width=450)
                    bild_angezeigt = True
                except Exception:
                    bild_angezeigt = False
            if not bild_angezeigt:
                st.markdown('<div class="ueber-mich-platzhalter">KW</div>', unsafe_allow_html=True)
        with col_text:
            st.markdown("### Über mich & das Projekt")
            st.markdown(
                "**Entwickelt von Kathrin Wozniak** – Architektin mit Fokus auf Bestandsmodernisierung und OpenBIM im Wohnungsbau.\n\n"
                "In der täglichen Planungspraxis liegen Bestandsgebäude heute oft als IFC-Modelle vor, "
                "doch für schnelle Wirtschaftlichkeits- und Sanierungsbetrachtungen fehlen häufig passgenaue, "
                "unkomplizierte Auswertungswerkzeuge ohne komplexe Spezialsoftware. "
                "Aus diesem Bedarf heraus – und im Rahmen meiner Weiterbildung im Bereich **BIM (Building Information Modeling)** – "
                "ist *bauteilauswertung-ifc* entstanden: ein praxisnahes Werkzeug aus der Architektur für die Architektur, "
                "das offene Gebäudedaten direkt im Browser nutzbar macht."
            )
    st.stop()


# ==============================================================================
# 5. FILTER, KATEGORISIERUNG & KENNZAHLEN
# ==============================================================================
df_waende, df_fenster, df_tueren, df_daecher, df_gelaender, df_entwaesserung = auswerten(datei.getvalue(), rechnen)
geschosse_waende = sorted(df_waende["Geschoss"].unique()) if not df_waende.empty else []

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

    st.subheader("Türauswertung (Suchbegriffe)")
    keller_kw = st.text_input("Kellergeschoss erkennen an", "keller, ug, untergeschoss, basement")
    dg_kw = st.text_input("Dachgeschoss erkennen an", "dachgeschoss, dg, spitzboden, attic")
    we_kw = st.text_input("Wohnungseingang erkennen an (Name/Typ)", "we, wohnungseingang, wohnungsabschluss")
    he_kw = st.text_input("Hauseingang erkennen an (Name/Typ)", "haustür, hauseingang, eingangstür")


def kategorie_wand(r):
    if not r.ist_aussen:
        return "Innenwand"
    if any(k.strip().lower() in r.Geschoss.lower() for k in fund_kw.split(",") if k.strip()):
        return "Fundament (unterirdisch)"
    if sock_kw.strip().lower() in r.Name.lower():
        return "Sockel (oberirdisch)"
    return "Außenwand Fassade" if r.Geschoss in dämm else "Außenwand sonstige"


def kategorie_tuer(r):
    """Ordnet jede Tür einer der 6 definierten Türarten zu."""
    geschoss_low = str(r.Geschoss).lower()
    name_typ_low = f"{r.Name} {r.Typ}".lower()

    ist_keller = any(k.strip().lower() in geschoss_low for k in keller_kw.split(",") if k.strip())
    ist_dg = any(k.strip().lower() in geschoss_low for k in dg_kw.split(",") if k.strip())
    ist_we = any(k.strip().lower() in name_typ_low for k in we_kw.split(",") if k.strip())
    ist_he = any(k.strip().lower() in name_typ_low for k in he_kw.split(",") if k.strip())

    # 1. Außentüren unterscheiden (Kelleraußentür vs. Hauseingangstür)
    if r.ist_aussen or ist_he:
        if ist_keller and not ist_he:
            return "Kelleraußentür"
        return "Hauseingangstür"

    # 2. Innentüren nach Funktion & Geschoss unterscheiden
    if ist_we:
        return "Wohnungseingangstür"
    if ist_keller:
        return "Kellertür"
    if ist_dg:
        return "Dachgeschosstür"
    return "Wohnungsinnentür"


df_waende["Kategorie"] = df_waende.apply(kategorie_wand, axis=1)
df_waende["F_ber"] = df_waende["Fläche"].where(df_waende["Herkunft"] == "berechnet", 0)
df_waende["ohne"] = df_waende["Fläche"].isna()

if not df_tueren.empty:
    df_tueren["Türart"] = df_tueren.apply(kategorie_tuer, axis=1)

fassade = df_waende.loc[df_waende.Kategorie == "Außenwand Fassade"]
sockel = df_waende.loc[(df_waende.Kategorie == "Sockel (oberirdisch)") & df_waende.Sockelmaterial.isin(sock_mat)]

dach_gesamt = df_daecher["Fläche (m²)"].sum() if not df_daecher.empty else 0.0
fenster_gesamt = df_fenster["Fläche (m²)"].sum() if not df_fenster.empty else 0.0
tueren_anzahl = len(df_tueren)
gelaender_gesamt = df_gelaender["Länge (m)"].sum() if not df_gelaender.empty else 0.0
rinnen_gesamt = (
    df_entwaesserung.loc[df_entwaesserung["Bauteilart"] == "Regenrinne", "Länge (m)"].sum()
    if not df_entwaesserung.empty else 0.0
)
fallrohre_gesamt = (
    df_entwaesserung.loc[df_entwaesserung["Bauteilart"] == "Fallrohr", "Länge (m)"].sum()
    if not df_entwaesserung.empty else 0.0
)

# Kennzahlen-Karten in 2 Reihen à 3 Spalten
k1, k2, k3 = st.columns(3)
k1.markdown(f'<div class="karte"><span>Fassadendämmung (netto)</span><b>{fassade.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="karte"><span>Sockeldämmung (netto)</span><b>{sockel.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="karte"><span>Dachfläche ({len(df_daecher)} Teilfl.)</span><b>{dach_gesamt:,.1f} m²</b></div>', unsafe_allow_html=True)

k4, k5, k6 = st.columns(3)
k4.markdown(f'<div class="karte"><span>Fenster ({len(df_fenster)} Stk.) & Türen ({tueren_anzahl} Stk.)</span><b>{fenster_gesamt:,.1f} m² | {tueren_anzahl} Türen</b></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="karte"><span>Geländer gesamt ({len(df_gelaender)} Stk.)</span><b>{gelaender_gesamt:,.1f} m</b></div>', unsafe_allow_html=True)
k6.markdown(f'<div class="karte"><span>Regenrinnen / Fallrohre</span><b>{rinnen_gesamt:,.1f} m / {fallrohre_gesamt:,.1f} m</b></div>', unsafe_allow_html=True)

fehlt = int(fassade.ohne.sum() + sockel.ohne.sum())
if fehlt:
    st.warning(f"{fehlt} Wände der Dämmflächen haben keine Fläche und fehlen in der Summe.")
if (
    (df_waende.Herkunft == "berechnet").any()
    or (not df_daecher.empty and (df_daecher.Herkunft == "berechnet").any())
    or (not df_gelaender.empty and (df_gelaender.Herkunft == "berechnet").any())
    or (not df_entwaesserung.empty and (df_entwaesserung.Herkunft == "berechnet").any())
):
    st.caption("Teile der Maße wurden aus der 3D-Geometrie berechnet (Spalte „Herkunft“).")


# ==============================================================================
# 6. TABELLEN & EXCEL-EXPORT
# ==============================================================================
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

if not df_fenster.empty:
    zus_fenster = df_fenster.groupby("Geschoss").agg(
        Anzahl=("Name", "size"),
        Fläche=("Fläche (m²)", "sum")
    ).reset_index()
    zus_fenster.loc[len(zus_fenster)] = ["Gesamt", zus_fenster["Anzahl"].sum(), zus_fenster["Fläche"].sum()]
else:
    zus_fenster = pd.DataFrame(columns=["Geschoss", "Anzahl", "Fläche (m²)"])

tuer_reihenfolge = [
    "Wohnungsinnentür",
    "Wohnungseingangstür",
    "Kellertür",
    "Dachgeschosstür",
    "Kelleraußentür",
    "Hauseingangstür"
]
if not df_tueren.empty:
    zus_tueren = df_tueren.groupby("Türart").agg(
        Anzahl=("Name", "size"),
        Fläche=("Fläche (m²)", "sum")
    ).reindex(tuer_reihenfolge, fill_value=0).reset_index()
    zus_tueren.loc[len(zus_tueren)] = ["Gesamt", zus_tueren["Anzahl"].sum(), zus_tueren["Fläche"].sum()]
    zus_tueren = zus_tueren.rename(columns={"Fläche": "Fläche (m²)"})
else:
    zus_tueren = pd.DataFrame(columns=["Türart", "Anzahl", "Fläche (m²)"])

if not df_daecher.empty:
    zus_daecher = df_daecher.groupby("Dachform").agg(
        Anzahl=("Name", "size"),
        Fläche=("Fläche (m²)", "sum")
    ).reset_index()
    zus_daecher.loc[len(zus_daecher)] = ["Gesamt", zus_daecher["Anzahl"].sum(), zus_daecher["Fläche"].sum()]
else:
    zus_daecher = pd.DataFrame(columns=["Dachform", "Anzahl", "Fläche (m²)"])

if not df_gelaender.empty:
    zus_gelaender = df_gelaender.groupby("Geländerart").agg(
        Anzahl=("Name", "size"),
        Länge=("Länge (m)", "sum")
    ).reset_index()
    zus_gelaender.loc[len(zus_gelaender)] = ["Gesamt", zus_gelaender["Anzahl"].sum(), zus_gelaender["Länge"].sum()]
    zus_gelaender = zus_gelaender.rename(columns={"Länge": "Länge (m)"})
else:
    zus_gelaender = pd.DataFrame(columns=["Geländerart", "Anzahl", "Länge (m)"])

if not df_entwaesserung.empty:
    zus_entwaesserung = df_entwaesserung.groupby("Bauteilart").agg(
        Anzahl=("Name", "size"),
        Länge=("Länge (m)", "sum")
    ).reset_index()
    zus_entwaesserung.loc[len(zus_entwaesserung)] = ["Gesamt", zus_entwaesserung["Anzahl"].sum(), zus_entwaesserung["Länge"].sum()]
    zus_entwaesserung = zus_entwaesserung.rename(columns={"Länge": "Länge (m)"})
else:
    zus_entwaesserung = pd.DataFrame(columns=["Bauteilart", "Anzahl", "Länge (m)"])

detail_waende = df_waende[["Name", "GlobalId", "Geschoss", "Kategorie", "Fläche", "Herkunft"]].rename(columns={"Fläche": "Fläche (m²)"})
detail_fenster = (
    df_fenster[["Name", "GlobalId", "Geschoss", "ist_aussen", "Breite (m)", "Höhe (m)", "Fläche (m²)"]]
    if not df_fenster.empty
    else pd.DataFrame(columns=["Name", "GlobalId", "Geschoss", "ist_aussen", "Breite (m)", "Höhe (m)", "Fläche (m²)"])
)
detail_tueren = (
    df_tueren[["Name", "Typ", "Mutterwand", "GlobalId", "Geschoss", "Türart", "ist_aussen", "Breite (m)", "Höhe (m)", "Fläche (m²)"]]
    if not df_tueren.empty
    else pd.DataFrame(columns=["Name", "Typ", "Mutterwand", "GlobalId", "Geschoss", "Türart", "ist_aussen", "Breite (m)", "Höhe (m)", "Fläche (m²)"])
)
detail_daecher = (
    df_daecher[["Name", "IFC-Klasse", "GlobalId", "Geschoss", "Dachform", "Neigung (°)", "Fläche (m²)", "Herkunft"]]
    if not df_daecher.empty
    else pd.DataFrame(columns=["Name", "IFC-Klasse", "GlobalId", "Geschoss", "Dachform", "Neigung (°)", "Fläche (m²)", "Herkunft"])
)
detail_gelaender = (
    df_gelaender[["Name", "GlobalId", "Geschoss", "Geländerart", "Länge (m)", "Herkunft"]]
    if not df_gelaender.empty
    else pd.DataFrame(columns=["Name", "GlobalId", "Geschoss", "Geländerart", "Länge (m)", "Herkunft"])
)
detail_entwaesserung = (
    df_entwaesserung[["Name", "Typ", "GlobalId", "Geschoss", "Bauteilart", "Länge (m)", "Herkunft"]]
    if not df_entwaesserung.empty
    else pd.DataFrame(columns=["Name", "Typ", "GlobalId", "Geschoss", "Bauteilart", "Länge (m)", "Herkunft"])
)

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

st.markdown("### Türen: Zusammenfassung nach Türart")
if not zus_tueren.empty:
    st.table(fmt(zus_tueren))
else:
    st.info("Keine Türen (IfcDoor) im Modell gefunden.")

st.markdown("### Dächer: Zusammenfassung")
if not zus_daecher.empty:
    st.table(fmt(zus_daecher))
else:
    st.info("Keine Dachflächen (IfcRoof / IfcSlab mit PredefinedType ROOF) im Modell gefunden.")

st.markdown("### Geländer: Zusammenfassung")
if not zus_gelaender.empty:
    st.table(fmt(zus_gelaender))
else:
    st.info("Keine Geländer (IfcRailing) im Modell gefunden.")

st.markdown("### Regenrinnen & Fallrohre: Zusammenfassung")
if not zus_entwaesserung.empty:
    st.table(fmt(zus_entwaesserung))
else:
    st.info("Keine Regenrinnen oder Fallrohre (IfcFlowSegment / IfcBuildingElementProxy) im Modell gefunden.")

with st.expander("Alle Wände einzeln"):
    st.dataframe(detail_waende, width="stretch", hide_index=True)
with st.expander("Alle Fenster einzeln"):
    st.dataframe(detail_fenster, width="stretch", hide_index=True)
with st.expander("Alle Türen einzeln"):
    st.dataframe(detail_tueren, width="stretch", hide_index=True)
with st.expander("Alle Dächer einzeln"):
    st.dataframe(detail_daecher, width="stretch", hide_index=True)
with st.expander("Alle Geländer einzeln"):
    st.dataframe(detail_gelaender, width="stretch", hide_index=True)
with st.expander("Alle Regenrinnen & Fallrohre einzeln"):
    st.dataframe(detail_entwaesserung, width="stretch", hide_index=True)

ergebnis = pd.DataFrame({
    "Ergebnis": [
        "Fassadendämmung netto (m²)",
        "Sockeldämmung netto (m²)",
        "Fensterfläche gesamt (m²)",
        "Türen gesamt (Anzahl)",
        "Dachfläche gesamt (m²)",
        "Geländer gesamt (m)",
        "Regenrinnen gesamt (m)",
        "Fallrohre gesamt (m)"
    ],
    "Wert": [
        round(fassade.Fläche.sum(), 1),
        round(sockel.Fläche.sum(), 1),
        round(fenster_gesamt, 1),
        int(tueren_anzahl),
        round(dach_gesamt, 1),
        round(gelaender_gesamt, 1),
        round(rinnen_gesamt, 1),
        round(fallrohre_gesamt, 1)
    ]
})

buf = io.BytesIO()
sheets = {
    "Ergebnis": ergebnis,
    "Wände Zusammenfassung": zus_waende,
    "Wände pro Geschoss": geo_waende,
    "Fenster Zusammenfassung": zus_fenster,
    "Türen Zusammenfassung": zus_tueren,
    "Dächer Zusammenfassung": zus_daecher,
    "Geländer Zusammenfassung": zus_gelaender,
    "Rinnen & Fallrohre": zus_entwaesserung,
    "Alle Wände": detail_waende,
    "Alle Fenster": detail_fenster,
    "Alle Türen": detail_tueren,
    "Alle Dächer": detail_daecher,
    "Alle Geländer": detail_gelaender,
    "Alle Rinnen & Fallrohre": detail_entwaesserung
}

with pd.ExcelWriter(buf, engine="openpyxl") as w:
    for name, d in sheets.items():
        d.to_excel(w, sheet_name=name, index=False)
    for ws in w.book.worksheets:
        for c in ws[1]:
            c.fill = PatternFill("solid", fgColor=HAUPT_BEIGE[1:])
            c.font = Font(color="FFFFFF", bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(60, max(len(str(c.value or "")) for c in col) + 3)

st.download_button(
    "Als Excel exportieren",
    buf.getvalue(),
    "bauteilauswertung-ifc.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
