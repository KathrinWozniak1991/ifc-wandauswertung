import io
import re
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element as el
import ifcopenshell.util.shape as shp
import ifcopenshell.util.unit as un
from openpyxl.styles import Font, PatternFill

# ==============================================================================
# 1. DARK-TECH DESIGN-KONFIGURATION, DATENSCHUTZ-SCHALTER & SEITENLAYOUT
# ==============================================================================
# Schalter für den Serverstandort-Hinweis (True = Serverstandort Deutschland aktiv)
HOSTING_IN_DEUTSCHLAND = True

# Futuristische OpenBIM Dark-Tech Farbpalette
BG_TIEFBLAU     = "#060E1B"   # Haupthintergrund (Deep Space Navy)
BG_PANEL        = "#0D1B30"   # Seitenleiste & Karten-Basis
AKZENT_CYAN     = "#00D4FF"   # Leuchtendes Cyan für Fokus, Rahmen & Kennzahlen
AKZENT_KOBALT   = "#2563EB"   # Technisches Kobaltblau
AKZENT_AMBER    = "#F59E0B"   # Signal-Amber/Gold für Kontraste & Warnungen
TEXT_HELL       = "#E6F1FF"   # Kühles Eisweiß für optimale Lesbarkeit
TEXT_GEDAEMPFT  = "#94A3B8"   # Blaugrau für Untertitel & Beschreibungen
EXCEL_HEADER    = "0D254C"    # Tiefblau für Excel-Tabellenköpfe

# Blickdichte Architektur- & CAD-Farben für die 3D-Gebäudevorschau (deckkraft = 1.00 = nicht durchsichtig)
BIM_FARBEN = {
    "Außenwände & Fassade":         {"farbe": "#E2E8F0", "deckkraft": 1.00, "standard_an": True},
    "Sockel (oberirdisch)":         {"farbe": "#2563EB", "deckkraft": 1.00, "standard_an": True},
    "Fundament & Erdreich-Sockel":  {"farbe": "#475569", "deckkraft": 1.00, "standard_an": True},
    "Geschossdecken":               {"farbe": "#94A3B8", "deckkraft": 1.00, "standard_an": True},
    "Innenwände":                   {"farbe": "#CBD5E1", "deckkraft": 1.00, "standard_an": False},
    "Dachflächen":                  {"farbe": "#C85A32", "deckkraft": 1.00, "standard_an": True},
    "Fenster":                      {"farbe": "#0284C7", "deckkraft": 1.00, "standard_an": True},
    "Türen":                        {"farbe": "#1E3A8A", "deckkraft": 1.00, "standard_an": True},
    "Geländer":                     {"farbe": "#D97706", "deckkraft": 1.00, "standard_an": True},
    "Regenrinnen & Fallrohre":      {"farbe": "#0891B2", "deckkraft": 1.00, "standard_an": True},
}

st.set_page_config(page_title="bauteilauswertung-ifc // OpenBIM Engine", layout="wide")

st.markdown(f"""<style>
/* Globales Dark-Tech-Layout & technisches Blueprint-Raster */
[data-testid="stAppViewContainer"] {{
    background-color: {BG_TIEFBLAU};
    background-image:
        radial-gradient(circle at 15% 15%, rgba(0, 212, 255, 0.08) 0%, transparent 40%),
        radial-gradient(circle at 85% 25%, rgba(29, 78, 216, 0.12) 0%, transparent 45%),
        linear-gradient(rgba(0, 212, 255, 0.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(0, 212, 255, 0.03) 1px, transparent 1px);
    background-size: 100% 100%, 100% 100%, 36px 36px, 36px 36px;
    color: {TEXT_HELL};
}}
[data-testid="stHeader"] {{
    background: rgba(6, 14, 27, 0.85);
    backdrop-filter: blur(8px);
}}
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #091529 0%, #0D1E38 100%);
    border-right: 1px solid rgba(0, 212, 255, 0.25);
}}
[data-testid="stSidebar"] * {{
    color: {TEXT_HELL};
}}

/* ============================================================================
   KONTRAST-FIX: UPLOAD-DROPZONE, UPLOAD-BUTTON, DATEINAME, SPINNER & EXPANDER
   ============================================================================ */
/* 1. Upload-Dropzone (Hauptseite, Seitenleiste & CSV-Uploader) */
[data-testid="stFileUploader"] section,
section[data-testid="stFileUploaderDropzone"] {{
    background-color: #0F2544 !important;
    border: 2px dashed {AKZENT_CYAN} !important;
    border-radius: 12px !important;
    padding: 16px 20px !important;
    color: #FFFFFF !important;
}}
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small {{
    color: #E6F1FF !important;
    opacity: 1 !important;
}}

/* 2. Innerer "Upload" / "Browse files"-Button im File-Uploader */
[data-testid="stFileUploader"] button,
section[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"],
[data-testid="stFileUploader"] [data-testid="stBaseButton-minimal"] {{
    background-color: #0284C7 !important;
    background-image: linear-gradient(90deg, #0369A1 0%, #0284C7 100%) !important;
    color: #FFFFFF !important;
    border: 1.5px solid #38BDF8 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    padding: 6px 16px !important;
    box-shadow: 0 0 12px rgba(0, 212, 255, 0.25) !important;
}}
[data-testid="stFileUploader"] button:hover,
section[data-testid="stFileUploaderDropzone"] button:hover {{
    background-color: #0EA5E9 !important;
    background-image: linear-gradient(90deg, #0284C7 0%, #0EA5E9 100%) !important;
    border-color: #FFFFFF !important;
}}
[data-testid="stFileUploader"] button *,
[data-testid="stFileUploader"] button svg {{
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
    stroke: #FFFFFF !important;
}}

/* 3. Zeile der bereits hochgeladenen Datei (Dateiname + Größe + Löschen-X) */
[data-testid="stFileUploaderFile"],
[data-testid="stFileUploader"] ul li,
[data-testid="stFileUploader"] ul li > div {{
    background-color: #132F57 !important;
    border: 1px solid {AKZENT_CYAN} !important;
    border-radius: 8px !important;
    color: #FFFFFF !important;
}}
[data-testid="stFileUploaderFile"] *,
[data-testid="stFileUploaderFileName"],
[data-testid="stFileUploader"] ul li * {{
    color: #FFFFFF !important;
    opacity: 1 !important;
}}
[data-testid="stFileUploaderFile"] svg,
[data-testid="stFileUploader"] ul li svg {{
    fill: {AKZENT_CYAN} !important;
    color: {AKZENT_CYAN} !important;
}}

/* 4. Lade-Spinner (st.spinner) */
[data-testid="stSpinner"],
[data-testid="stSpinner"] > div,
div.stSpinner > div {{
    background-color: #0F2544 !important;
    color: #FFFFFF !important;
    border: 1.5px solid {AKZENT_CYAN} !important;
    border-radius: 10px !important;
    padding: 10px 16px !important;
}}
[data-testid="stSpinner"] *,
div.stSpinner * {{
    color: #FFFFFF !important;
    font-weight: 600 !important;
    opacity: 1 !important;
}}

/* 5. Ausklappbare Boxen (st.expander) */
[data-testid="stExpander"],
[data-testid="stExpander"] details {{
    background-color: #0D1B30 !important;
    border: 1px solid rgba(0, 212, 255, 0.45) !important;
    border-radius: 10px !important;
}}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] details > summary {{
    background-color: #102544 !important;
    color: #FFFFFF !important;
    border-radius: 9px !important;
    padding: 12px 16px !important;
}}
[data-testid="stExpander"] summary *,
[data-testid="stExpander"] details > summary * {{
    color: #FFFFFF !important;
    font-weight: 600 !important;
    opacity: 1 !important;
}}
[data-testid="stExpander"] summary:hover {{
    background-color: #16335E !important;
}}

/* Eingabefelder & Multiselects im Dark-Tech-Design */
[data-baseweb="input"] > div,
[data-baseweb="select"] > div {{
    background-color: #102544 !important;
    border-color: rgba(0, 212, 255, 0.35) !important;
    color: #FFFFFF !important;
}}
[data-baseweb="input"] input,
[data-baseweb="select"] * {{
    color: #FFFFFF !important;
}}

/* Prominente Upload-Einladungs-Box auf der Startseite */
.upload-hero {{
    background: linear-gradient(135deg, rgba(13, 34, 66, 0.95) 0%, rgba(10, 56, 94, 0.92) 100%);
    border: 2px solid {AKZENT_CYAN};
    border-radius: 14px;
    padding: 20px 24px 12px 24px;
    margin-bottom: 18px;
    box-shadow: 0 0 28px rgba(0, 212, 255, 0.20);
}}
.upload-hero h3 {{
    margin: 0 0 4px 0 !important;
    color: #FFFFFF !important;
    font-size: 1.25rem !important;
}}
.upload-hero p {{
    margin: 0 0 12px 0 !important;
    color: #CBD5E1 !important;
    font-size: 0.95rem !important;
}}

/* Futuristischer Hero-Header */
.kopf {{
    background: linear-gradient(135deg, #08162C 0%, #0F2B57 55%, #0A4D7C 100%);
    color: #FFFFFF;
    padding: 24px 28px;
    border-radius: 12px;
    border: 1px solid rgba(0, 212, 255, 0.45);
    border-left: 6px solid {AKZENT_CYAN};
    box-shadow: 0 8px 32px rgba(0, 212, 255, 0.14);
    margin-bottom: 16px;
}}
.kopf-tag {{
    display: inline-block;
    font-family: monospace;
    font-size: 0.78rem;
    letter-spacing: 0.08em;
    color: {AKZENT_CYAN};
    background: rgba(0, 212, 255, 0.12);
    border: 1px solid rgba(0, 212, 255, 0.35);
    padding: 3px 10px;
    border-radius: 4px;
    margin-bottom: 8px;
}}
.kopf h1 {{
    color: #FFFFFF;
    margin: 0;
    font-size: 1.85rem;
    letter-spacing: 0.01em;
    text-shadow: 0 0 18px rgba(0, 212, 255, 0.35);
}}
.kopf p {{
    color: #CBD5E1;
    margin: 6px 0 0 0;
    font-size: 1rem;
}}

/* Datenschutz- & Sicherheits-Banner */
.trust-banner {{
    background: linear-gradient(90deg, rgba(10, 38, 46, 0.90) 0%, rgba(13, 33, 58, 0.90) 100%);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(16, 185, 129, 0.45);
    border-left: 5px solid #10B981;
    padding: 13px 18px;
    border-radius: 10px;
    margin-bottom: 12px;
    color: #E2F8F0;
    font-size: 0.93rem;
    line-height: 1.45;
    box-shadow: 0 4px 20px rgba(16, 185, 129, 0.08);
}}
.trust-banner strong {{
    color: #6EE7B7;
}}
.trust-badges {{
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 8px;
    font-weight: 600;
    font-size: 0.84rem;
}}
.trust-badge-item {{
    background: rgba(16, 185, 129, 0.16);
    color: #A7F3D0;
    padding: 4px 11px;
    border-radius: 5px;
    border: 1px solid rgba(16, 185, 129, 0.40);
}}
.sidebar-trust {{
    background: rgba(10, 38, 46, 0.82);
    border: 1px solid rgba(16, 185, 129, 0.40);
    border-left: 4px solid #10B981;
    padding: 10px 12px;
    border-radius: 8px;
    font-size: 0.82rem;
    color: #D1FAE5;
    margin-top: 8px;
    margin-bottom: 14px;
    line-height: 1.4;
}}

/* Glassmorphism Kennzahlen-Karten */
.karte {{
    background: linear-gradient(145deg, rgba(15, 34, 64, 0.84), rgba(10, 24, 46, 0.94));
    backdrop-filter: blur(12px);
    border: 1px solid rgba(0, 212, 255, 0.30);
    border-left: 4px solid {AKZENT_CYAN};
    padding: 15px 18px;
    border-radius: 10px;
    margin-bottom: 12px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.32);
    transition: transform 0.18s ease, border-color 0.18s ease, box-shadow 0.18s ease;
}}
.karte:hover {{
    transform: translateY(-2px);
    border-color: rgba(0, 212, 255, 0.78);
    box-shadow: 0 6px 24px rgba(0, 212, 255, 0.20);
}}
.karte span {{
    color: {TEXT_GEDAEMPFT};
    font-size: 0.84rem;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}}
.karte b {{
    display: block;
    color: #FFFFFF;
    font-size: 1.50rem;
    margin-top: 3px;
    text-shadow: 0 0 12px rgba(0, 212, 255, 0.25);
}}

/* Glassmorphism Feature-Boxen auf der Startseite */
.info-box {{
    background: linear-gradient(160deg, rgba(16, 36, 68, 0.78), rgba(9, 21, 41, 0.92));
    backdrop-filter: blur(12px);
    border: 1px solid rgba(0, 212, 255, 0.28);
    border-top: 3px solid {AKZENT_CYAN};
    padding: 18px 20px;
    border-radius: 10px;
    height: 100%;
    box-shadow: 0 6px 24px rgba(0, 0, 0, 0.35);
    transition: border-color 0.2s ease, box-shadow 0.2s ease;
}}
.info-box:hover {{
    border-color: rgba(0, 212, 255, 0.68);
    box-shadow: 0 8px 28px rgba(0, 212, 255, 0.16);
}}
.info-box .modul-nr {{
    font-family: monospace;
    font-size: 0.78rem;
    color: {AKZENT_CYAN};
    letter-spacing: 0.08em;
    margin-bottom: 4px;
}}
.info-box h4 {{
    color: #FFFFFF;
    margin-top: 0;
    margin-bottom: 8px;
    font-size: 1.08rem;
}}
.info-box p {{
    color: #CBD5E1;
    font-size: 0.93rem;
    line-height: 1.5;
    margin: 0;
}}

/* Profilbild-Platzhalter im Dark-Tech-Stil */
.ueber-mich-platzhalter {{
    width: 145px;
    height: 145px;
    border-radius: 50%;
    border: 3px solid {AKZENT_CYAN};
    box-shadow: 0 0 22px rgba(0, 212, 255, 0.35);
    background: linear-gradient(135deg, #0D254C, #0077B6);
    color: #FFFFFF;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 2rem;
    font-weight: bold;
    flex-shrink: 0;
}}

/* Tabellen, Überschriften & Buttons */
[data-testid="stTable"] table {{
    background-color: rgba(13, 27, 48, 0.88) !important;
    color: {TEXT_HELL} !important;
    border: 1px solid rgba(0, 212, 255, 0.25) !important;
}}
[data-testid="stTable"] th {{
    background: linear-gradient(90deg, #0D254C, #13386E) !important;
    color: {AKZENT_CYAN} !important;
    border-bottom: 1px solid rgba(0, 212, 255, 0.4) !important;
}}
[data-testid="stTable"] td {{
    color: {TEXT_HELL} !important;
    border-bottom: 1px solid rgba(148, 163, 184, 0.14) !important;
}}
h1, h2, h3, h4 {{
    color: #F8FAFC !important;
}}
p, li, label, .stMarkdown {{
    color: {TEXT_HELL};
}}
.stDownloadButton button, .stButton button {{
    background: linear-gradient(90deg, #0E3A6E 0%, #0284C7 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(0, 212, 255, 0.55) !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    box-shadow: 0 0 15px rgba(0, 212, 255, 0.18) !important;
    transition: all 0.2s ease !important;
}}
.stDownloadButton button:hover, .stButton button:hover {{
    border-color: #FFFFFF !important;
    box-shadow: 0 0 22px rgba(0, 212, 255, 0.45) !important;
}}
</style>""", unsafe_allow_html=True)

st.markdown(
    '<div class="kopf">'
    '<div class="kopf-tag">OPENBIM ANALYTICS ENGINE // IFC 2X3 &amp; IFC 4 // DIN 276</div>'
    '<h1>bauteilauswertung-ifc</h1>'
    '<p>Automatisierte 3D-Geometrie-, Mengen- und Kostengruppenanalyse für Bestandsgebäude</p>'
    '</div>',
    unsafe_allow_html=True
)

# ==============================================================================
# 1B. SICHTBARES DATENSCHUTZ- & SICHERHEITS-BANNER + DETAIL-BOX
# ==============================================================================
server_badge = (
    '<span class="trust-badge-item">Serverstandort Deutschland (DSGVO)</span>'
    if HOSTING_IN_DEUTSCHLAND
    else '<span class="trust-badge-item">DSGVO-konforme Verarbeitung</span>'
)

st.markdown(
    f'<div class="trust-banner">'
    f'<strong>Zero-Storage &amp; Datenschutz-Garantie:</strong> '
    f'Ihre IFC-Gebäudemodelle und Auswertungsdaten werden <strong>nicht dauerhaft gespeichert</strong>, '
    f'<strong>nicht an Dritte weitergegeben</strong> und <strong>nicht für KI-Training verwendet</strong>.'
    f'<div class="trust-badges">'
    f'<span class="trust-badge-item">Keine Speicherung von IFC-Dateien</span>'
    f'<span class="trust-badge-item">Flüchtige RAM-Auswertung</span>'
    f'<span class="trust-badge-item">Keine Weitergabe an Dritte</span>'
    f'{server_badge}'
    f'</div>'
    f'</div>',
    unsafe_allow_html=True
)

with st.expander("Technische Details zu Datenschutz & Datensicherheit (für Auftraggeber & IT-Prüfung)", expanded=False):
    st.markdown(
        "**1. Keine dauerhafte Dateispeicherung:**\n"
        "- Beim Hochladen wird Ihre `.ifc`-Datei ausschließlich für den Bruchteil einer Sekunde zum Einlesen der Geometrie "
        "in eine isolierte, temporäre Systemdatei geladen und unmittelbar nach dem Einlesen (`finally: unlink()`) automatisch wieder gelöscht.\n"
        "- Es existieren weder Datenbanken noch Cloud-Speicher-Buckets, in denen hochgeladene Modelle archiviert werden.\n\n"
        "**2. Isolierte Verarbeitung in Ihrer Browser-Sitzung:**\n"
        "- Alle Berechnungen (Mengenermittlung, 3D-Vorschau, DIN 276-Zuordnung und Excel-Generierung) erfolgen flüchtig im Arbeitsspeicher (RAM) Ihrer aktuellen Sitzung.\n"
        "- Andere Nutzer haben zu keinem Zeitpunkt Zugriff auf Ihre Sitzungsdaten. Beim Schließen des Browser-Tabs oder über den Button "
        "*„Sitzung & Arbeitsspeicher leeren“* werden alle Zwischendaten vollständig verworfen.\n\n"
        "**3. Keine Weitergabe & kein KI-Training:**\n"
        "- Ihre Gebäudemodelle, Grundrisse, Bauteileigenschaften und Preisangaben werden weder an externe Dienste übertragen noch ausgewertet oder für das Training von KI-Modellen genutzt.\n\n"
        "**4. Verschlüsselte Übertragung & Hosting:**\n"
        "- Die Datenübertragung zwischen Ihrem Browser und dem Server erfolgt durchgehend TLS/HTTPS-verschlüsselt. "
        + ("Der Betrieb erfolgt auf Serverinfrastruktur in Deutschland gemäß DSGVO." if HOSTING_IN_DEUTSCHLAND else "Die Verarbeitung erfolgt streng zweckgebunden gemäß DSGVO.")
    )


# ==============================================================================
# 2. HILFSFUNKTIONEN (3D-GEOMETRIE-FALLBACKS & 3D-MESH-EXTRAKTION)
# ==============================================================================
def berechne_wand(wand, settings, skala: float):
    """Berechnet die fehlende Seitenfläche einer Wand über Volumen / Wanddicke."""
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
        volumen = shp.get_volume(ifcopenshell.geom.create_shape(settings, wand).geometry)
        return volumen / d
    except Exception:
        return None


def berechne_horizontale_flaeche(bauteil, settings):
    """Fallback: Ermittelt horizontale bzw. geneigte Oberseite und mittlere Neigung aus der 3D-Geometrie (für Dächer & Decken)."""
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
    """Schätzt die Bauteillänge aus der 3D-Ausdehnung (Bounding Box) in Metern."""
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


def _hole_suchtext_bauteil(bauteil) -> str:
    """
    Fasst Instanzname, Bauteiltyp (IfcTypeObject) und Materialnamen eines IFC-Elements
    in einem kleingeschriebenen Suchstring zusammen.
    """
    teile = [bauteil.Name or ""]
    try:
        typ_obj = el.get_type(bauteil)
        if typ_obj and typ_obj.Name:
            teile.append(typ_obj.Name)
    except Exception:
        pass
    try:
        mat = el.get_material(bauteil, should_skip_usage=True)
        if mat:
            if hasattr(mat, "Name") and mat.Name:
                teile.append(mat.Name)
            if mat.is_a("IfcMaterialLayerSet"):
                teile.extend(
                    s.Material.Name for s in mat.MaterialLayers
                    if getattr(s, "Material", None) and s.Material.Name
                )
    except Exception:
        pass
    return " ".join(teile).lower()


@st.cache_data(show_spinner="3D-Gebäudemodell wird erstellt …")
def extrahiere_3d_meshes(daten: bytes, max_dreiecke_gesamt: int = 180000):
    """
    Erzeugt pro BIM-Bauteilgruppe ein kombiniertes Dreiecksnetz (Vertices + Faces)
    in Weltkoordinaten für die interaktive Plotly-3D-Vorschau.
    """
    with tempfile.NamedTemporaryFile(suffix=".ifc", delete=False) as f:
        f.write(daten)
        tmp_pfad = Path(f.name)

    try:
        m = ifcopenshell.open(str(tmp_pfad))
    finally:
        tmp_pfad.unlink(missing_ok=True)

    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)

    gruppen_puffer = {
        k: {"verts": [], "faces": [], "offset": 0, "bauteile": 0}
        for k in BIM_FARBEN
    }
    dreiecke_gesamt = 0
    gekurzt = False

    def fuege_bauteil_hinzu(bauteil, gruppe_name: str, z_pruefung_erdreich: bool = False):
        nonlocal dreiecke_gesamt, gekurzt
        if dreiecke_gesamt >= max_dreiecke_gesamt:
            gekurzt = True
            return
        try:
            shape = ifcopenshell.geom.create_shape(settings, bauteil)
            geom = shape.geometry
            verts = np.array(geom.verts, dtype=np.float32).reshape(-1, 3)
            faces = np.array(geom.faces, dtype=np.int32).reshape(-1, 3)
            if len(verts) == 0 or len(faces) == 0:
                return

            if z_pruefung_erdreich and gruppe_name in {"Außenwände & Fassade", "Innenwände"}:
                z_max = float(verts[:, 2].max())
                if z_max <= -0.10:
                    gruppe_name = "Fundament & Erdreich-Sockel"

            puffer = gruppen_puffer[gruppe_name]
            puffer["verts"].append(verts)
            puffer["faces"].append(faces + puffer["offset"])
            puffer["offset"] += len(verts)
            puffer["bauteile"] += 1
            dreiecke_gesamt += len(faces)
        except Exception:
            pass

    erdreich_gesch_kw = ["foundation", "fundament", "keller", "ug", "untergeschoss", "basement", "erdreich", "gründung"]
    fundament_kw = ["fundament", "streifenfundament", "perimeter", "erdreich", "kelleraußen", "kelleraussen", "unterirdisch", "gründung"]
    sockel_kw = ["sockel", "plinth", "spritzwasser"]

    for w in m.by_type("IfcWall"):
        ps = el.get_psets(w)
        suchtext = _hole_suchtext_bauteil(w)
        g = el.get_container(w)
        geschoss_name = (g.Name or "").lower() if g else ""

        ist_aussen = ps.get("Pset_WallCommon", {}).get("IsExternal") is True or any(
            k in suchtext for k in ["fassade", "aussen", "außen", "exterior", "sockel", "perimeter", "fundament", "kelleraußen"]
        )
        ist_fundament_text = any(k in suchtext for k in fundament_kw) or any(
            k in geschoss_name for k in ["foundation", "fundament", "gründung"]
        )
        ist_sockel_text = any(k in suchtext for k in sockel_kw)
        ist_im_kellergeschoss = any(k in geschoss_name for k in erdreich_gesch_kw)

        if ist_fundament_text:
            fuege_bauteil_hinzu(w, "Fundament & Erdreich-Sockel")
        elif ist_sockel_text:
            fuege_bauteil_hinzu(w, "Sockel (oberirdisch)")
        elif ist_im_kellergeschoss:
            fuege_bauteil_hinzu(w, "Fundament & Erdreich-Sockel")
        elif ist_aussen:
            fuege_bauteil_hinzu(w, "Außenwände & Fassade", z_pruefung_erdreich=True)
        else:
            fuege_bauteil_hinzu(w, "Innenwände", z_pruefung_erdreich=True)

    for klasse in ["IfcFooting", "IfcBeam", "IfcMember", "IfcColumn"]:
        try:
            for elem in m.by_type(klasse):
                suchtext = _hole_suchtext_bauteil(elem)
                g = el.get_container(elem)
                geschoss_name = (g.Name or "").lower() if g else ""
                if (
                    klasse == "IfcFooting"
                    or any(k in f"{suchtext} {geschoss_name}" for k in ["fundament", "streifen", "gründung", "footing", "sohle"])
                    or (klasse == "IfcColumn" and any(k in geschoss_name for k in erdreich_gesch_kw))
                ):
                    fuege_bauteil_hinzu(elem, "Fundament & Erdreich-Sockel")
                elif "sockel" in suchtext:
                    fuege_bauteil_hinzu(elem, "Sockel (oberirdisch)")
        except Exception:
            pass

    dach_liste = []
    for r in m.by_type("IfcRoof"):
        teile = el.get_parts(r)
        dach_liste.extend(teile if teile else [r])

    for s in m.by_type("IfcSlab"):
        predef = (getattr(s, "PredefinedType", None) or "").upper()
        suchtext = _hole_suchtext_bauteil(s)
        g = el.get_container(s)
        geschoss_name = (g.Name or "").lower() if g else ""

        if predef == "ROOF" or s in dach_liste:
            if s not in dach_liste:
                dach_liste.append(s)
        elif (
            predef == "BASESLAB"
            or any(k in suchtext for k in ["bodenplatte", "fundament", "sohle", "gründung", "kellerboden"])
            or any(k in geschoss_name for k in ["foundation", "fundament", "gründung"])
        ):
            fuege_bauteil_hinzu(s, "Fundament & Erdreich-Sockel")
        else:
            fuege_bauteil_hinzu(s, "Geschossdecken")

    for d in dach_liste:
        fuege_bauteil_hinzu(d, "Dachflächen")
    for f in m.by_type("IfcWindow"):
        fuege_bauteil_hinzu(f, "Fenster")
    for t in m.by_type("IfcDoor"):
        fuege_bauteil_hinzu(t, "Türen")
    for r in m.by_type("IfcRailing"):
        fuege_bauteil_hinzu(r, "Geländer")

    for klasse in ["IfcFlowSegment", "IfcBuildingElementProxy"]:
        try:
            for p in m.by_type(klasse):
                predef = (getattr(p, "PredefinedType", None) or "").upper()
                suchtext = f"{_hole_suchtext_bauteil(p)} {predef}".lower()
                if predef in {"GUTTER", "DOWNPIPE"} or any(
                    k in suchtext for k in ["rinne", "dachrinne", "gutter", "fallrohr", "regenrohr", "downpipe"]
                ):
                    fuege_bauteil_hinzu(p, "Regenrinnen & Fallrohre")
                elif any(k in suchtext for k in ["fundament", "streifen", "gründung", "footing"]):
                    fuege_bauteil_hinzu(p, "Fundament & Erdreich-Sockel")
                elif "sockel" in suchtext:
                    fuege_bauteil_hinzu(p, "Sockel (oberirdisch)")
        except Exception:
            pass

    ergebnis_meshes = {}
    for gruppe_name, puffer in gruppen_puffer.items():
        if puffer["verts"]:
            ergebnis_meshes[gruppe_name] = {
                "verts": np.vstack(puffer["verts"]),
                "faces": np.vstack(puffer["faces"]),
                "bauteile": puffer["bauteile"],
            }

    return ergebnis_meshes, dreiecke_gesamt, gekurzt


# ==============================================================================
# 3. HAUPTAUSWERTUNG DER IFC-DATEI
# ==============================================================================
@st.cache_data(show_spinner="IFC-Modell wird analysiert … (Wände, Decken, Fenster, Türen, Dächer & Geländer)")
def auswerten(daten: bytes, rechnen: bool):
    """
    Liest das hochgeladene IFC-Modell flüchtig ein und löscht die temporäre
    Arbeitsdatei unmittelbar nach dem Öffnen im finally-Block.
    """
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
        suchtext_wand = _hole_suchtext_bauteil(w)
        ist_aussen = ps.get("Pset_WallCommon", {}).get("IsExternal") is True
        if not ist_aussen and any(k in suchtext_wand for k in ["fassade", "aussen", "außen", "exterior", "sockel", "perimeter", "fundament"]):
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
        if (getattr(s, "PredefinedType", None) or "").upper() == "ROOF" and s not in dach_kandidaten:
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
            fl_geo, neig_geo = berechne_horizontale_flaeche(d, settings)
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

    # --- 4B. Geschossdecken & Bodenplatten auswerten (IfcSlab ohne Dächer) ---
    zeilen_decken = []
    for s in m.by_type("IfcSlab"):
        predef = (getattr(s, "PredefinedType", None) or "").upper()
        if predef == "ROOF" or s in dach_kandidaten:
            continue

        ps = el.get_psets(s)
        q = ps.get("Qto_SlabBaseQuantities") or ps.get("BaseQuantities") or {}
        fl = q.get("NetArea") if isinstance(q.get("NetArea"), (int, float)) else q.get("GrossArea")
        herkunft = "Modell" if isinstance(fl, (int, float)) else "fehlt"

        if fl is None and rechnen:
            fl_geo, _ = berechne_horizontale_flaeche(s, settings)
            if fl_geo is not None:
                fl = fl_geo
                herkunft = "berechnet"

        g = el.get_container(s)
        geschoss_name = g.Name if g else "ohne Geschoss"
        suchtext_decke = f"{_hole_suchtext_bauteil(s)} {geschoss_name} {predef}".lower()

        if (
            predef == "BASESLAB"
            or any(k in suchtext_decke for k in ["bodenplatte", "fundament", "sohle", "gründung", "kellerboden", "foundation", "basement", "keller", "ug"])
        ):
            deckenart = "Bodenplatte / Kellerdecke"
        elif any(k in suchtext_decke for k in ["oberste", "dachboden", "spitzboden", "kehlbalken", "attic", "dachgeschoss", "dg"]):
            deckenart = "Oberste Geschossdecke"
        else:
            deckenart = "Geschossdecke"

        zeilen_decken.append({
            "Name": s.Name or "Decke",
            "GlobalId": s.GlobalId,
            "Geschoss": geschoss_name,
            "Deckenart": deckenart,
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

    # --- 6. Regenrinnen & Fallrohre auswerten ---
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
        pd.DataFrame(zeilen_decken),
        pd.DataFrame(zeilen_gelaender),
        pd.DataFrame(zeilen_entwaesserung)
    )


# ==============================================================================
# 4. UPLOAD-BEREICH (PROMINENT + SEITENLEISTE) & STARTSEITE
# ==============================================================================
with st.sidebar:
    st.markdown(f"<h3 style='color:{AKZENT_CYAN};margin-bottom:4px;'>CONTROL PANEL</h3>", unsafe_allow_html=True)
    datei_sidebar = st.file_uploader("IFC-Modell hochladen (.ifc)", type="ifc", key="ifc_upload_sidebar")
    st.markdown(
        '<div class="sidebar-trust">'
        '<strong>100 % vertraulich:</strong> Ihre IFC-Datei wird ausschließlich flüchtig im Arbeitsspeicher '
        'ausgewertet, nicht gespeichert und nicht an Dritte weitergegeben.'
        '</div>',
        unsafe_allow_html=True
    )
    rechnen = st.checkbox("Fehlende Maße aus 3D-Geometrie berechnen", value=True)

datei_haupt = None
if datei_sidebar is None:
    st.markdown(
        '<div class="upload-hero">'
        '<h3>IFC-Gebäudemodell hochladen &amp; Auswertung starten</h3>'
        '<p>Ziehen Sie Ihre <strong>.ifc-Datei</strong> direkt in das Feld unten oder klicken Sie auf '
        '<strong>Upload / Browse files</strong> (alternativ auch links im Control Panel möglich).</p>'
        '</div>',
        unsafe_allow_html=True
    )
    datei_haupt = st.file_uploader(
        "IFC-Datei auswählen (.ifc)",
        type="ifc",
        key="ifc_upload_main",
        label_visibility="collapsed"
    )

datei = datei_sidebar if datei_sidebar is not None else datei_haupt

if datei is not None:
    with st.sidebar:
        if st.button("Sitzung & Arbeitsspeicher leeren", use_container_width=True, help="Löscht alle zwischengespeicherten Auswertungsdaten sofort aus dem RAM."):
            st.cache_data.clear()
            st.rerun()

if not datei:
    st.markdown("### Willkommen im OpenBIM-Auswertungs-Cockpit")
    st.write(
        "Diese Web-Anwendung analysiert digitale Gebäudemodelle im offenen **IFC-Format (OpenBIM)** "
        "vollautomatisch in Echtzeit. Entwickelt für Bestandsgebäude im Wohnungsbau, liefert das System "
        "in Sekunden belastbare Mengen, 3D-Ansichten und **DIN 276-Kostengruppen** für "
        "Modernisierungs- und Sanierungsprojekte."
    )

    v1, v2, v3, v4 = st.columns(4)
    with v1:
        st.markdown(
            '<div class="info-box">'
            '<div class="modul-nr">MODUL 01 // QUANTITY TAKE-OFF</div>'
            '<h4>Automatisierte Mengen</h4>'
            '<p>Fassaden-, Sockel-, Decken-, Fenster-, Tür- und Dachflächen sowie Geländer- und Rinnenlängen '
            'werden direkt aus der IFC-Datenstruktur extrahiert – ohne manuelles Nachmessen.</p>'
            '</div>',
            unsafe_allow_html=True
        )
    with v2:
        st.markdown(
            '<div class="info-box">'
            '<div class="modul-nr">MODUL 02 // 3D GEOMETRY ENGINE</div>'
            '<h4>3D-CAD-Check &amp; Fallback</h4>'
            '<p>Fehlen im Bestandsmodell Flächen- oder Längenparameter, berechnet die Geometrie-Engine '
            'diese aus den 3D-Körpern nach und kennzeichnet die Datenherkunft transparent.</p>'
            '</div>',
            unsafe_allow_html=True
        )
    with v3:
        st.markdown(
            '<div class="info-box">'
            '<div class="modul-nr">MODUL 03 // DIN 276 &amp; EXPORT</div>'
            '<h4>Kostengruppen &amp; Excel</h4>'
            '<p>Automatische Strukturierung nach DIN 276 (2. Ebene), Eingabe eigener Einheitspreise '
            'oder CSV-Preislisten sowie strukturierter Mehrblatt-Excel-Export.</p>'
            '</div>',
            unsafe_allow_html=True
        )
    with v4:
        st.markdown(
            '<div class="info-box">'
            '<div class="modul-nr">MODUL 04 // ZERO-STORAGE SECURITY</div>'
            '<h4>Maximale Datensicherheit</h4>'
            '<p>Keine Archivierung auf Servern, keine Weitergabe, kein KI-Training: '
            'Ihr Modell wird flüchtig im Arbeitsspeicher auf '
            + ('deutschen Servern ' if HOSTING_IN_DEUTSCHLAND else '')
            + 'verarbeitet und sofort verworfen.</p>'
            '</div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### Workflow in drei Schritten")
    st.markdown(
        "1. **IFC-Datei laden:** Wählen Sie oben im Upload-Feld (oder links im Control Panel) ein `.ifc`-Gebäudemodell aus.\n"
        "2. **Geschosse & eigene Preise konfigurieren:** Prüfen Sie die Geschossfilter in der Seitenleiste und hinterlegen Sie bei Bedarf eigene Einheitspreise nach DIN 276.\n"
        "3. **3D-Modell prüfen & Excel exportieren:** Kontrollieren Sie die Bauteile im interaktiven 3D-Viewport, laden Sie den **Excel-Bericht** herunter und leeren Sie bei Bedarf den Arbeitsspeicher."
    )

    # --- Bereich "Über mich & das Projekt" ---
    st.markdown("---")
    basis_ordner = Path(__file__).resolve().parent
    profil_pfad = next(
        (p for p in basis_ordner.iterdir() if p.is_file() and "profil" in p.stem.lower() and p.suffix.lower() in {".jpg", ".jpeg", ".png"}),
        None
    )

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
        border: 3px solid {AKZENT_CYAN};
        box-shadow: 0 0 22px rgba(0, 212, 255, 0.35);
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
ifc_bytes = datei.getvalue()
df_waende, df_fenster, df_tueren, df_daecher, df_decken, df_gelaender, df_entwaesserung = auswerten(ifc_bytes, rechnen)
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

    if r.ist_aussen or ist_he:
        if ist_keller and not ist_he:
            return "Kelleraußentür"
        return "Hauseingangstür"

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
decken_gesamt = df_decken["Fläche (m²)"].sum() if not df_decken.empty else 0.0
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

# Kennzahlen-Karten in 2 Reihen à 3 Spalten (inkl. Geschossdecken)
k1, k2, k3 = st.columns(3)
k1.markdown(f'<div class="karte"><span>Fassadendämmung (netto)</span><b>{fassade.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="karte"><span>Sockeldämmung (netto)</span><b>{sockel.Fläche.sum():,.1f} m²</b></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="karte"><span>Dach ({len(df_daecher)}) &amp; Decken ({len(df_decken)})</span><b>{dach_gesamt:,.1f} m² | {decken_gesamt:,.1f} m²</b></div>', unsafe_allow_html=True)

k4, k5, k6 = st.columns(3)
k4.markdown(f'<div class="karte"><span>Fenster ({len(df_fenster)} Stk.) &amp; Türen ({tueren_anzahl} Stk.)</span><b>{fenster_gesamt:,.1f} m² | {tueren_anzahl} Türen</b></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="karte"><span>Geländer gesamt ({len(df_gelaender)} Stk.)</span><b>{gelaender_gesamt:,.1f} m</b></div>', unsafe_allow_html=True)
k6.markdown(f'<div class="karte"><span>Regenrinnen / Fallrohre</span><b>{rinnen_gesamt:,.1f} m / {fallrohre_gesamt:,.1f} m</b></div>', unsafe_allow_html=True)

fehlt = int(fassade.ohne.sum() + sockel.ohne.sum())
if fehlt:
    st.warning(f"{fehlt} Wände der Dämmflächen haben keine Fläche und fehlen in der Summe.")
if (
    (df_waende.Herkunft == "berechnet").any()
    or (not df_daecher.empty and (df_daecher.Herkunft == "berechnet").any())
    or (not df_decken.empty and (df_decken.Herkunft == "berechnet").any())
    or (not df_gelaender.empty and (df_gelaender.Herkunft == "berechnet").any())
    or (not df_entwaesserung.empty and (df_entwaesserung.Herkunft == "berechnet").any())
):
    st.caption("Teile der Maße wurden aus der 3D-Geometrie berechnet (Spalte „Herkunft“).")


# ==============================================================================
# 5B. INTERAKTIVE 3D-GEBÄUDEVORSCHAU (HELLER STUDIO-VIEWPORT & BLICKDICHTES HAUS)
# ==============================================================================
with st.expander("Interaktive 3D-Gebäudevorschau (OpenBIM-Modell)", expanded=True):
    meshes_3d, dreiecke_anzahl, modell_gekurzt = extrahiere_3d_meshes(ifc_bytes)

    if not meshes_3d:
        st.info("Für dieses IFC-Modell konnten keine darstellbaren 3D-Körper erzeugt werden.")
    else:
        st.caption(
            "Steuerung: Linke Maustaste = Drehen | Rechte Maustaste = Verschieben | Mausrad = Zoomen. "
            "Bauteilgruppen im Layer-Manager ein- oder ausblenden:"
        )

        gruppen_namen = list(BIM_FARBEN.keys())
        aktive_gruppen = {}
        cb_spalten = st.columns(3)
        for idx, gruppe in enumerate(gruppen_namen):
            anzahl_bt = meshes_3d[gruppe]["bauteile"] if gruppe in meshes_3d else 0
            with cb_spalten[idx % 3]:
                aktive_gruppen[gruppe] = st.checkbox(
                    f"{gruppe} ({anzahl_bt})",
                    value=BIM_FARBEN[gruppe]["standard_an"] and (anzahl_bt > 0),
                    disabled=(anzahl_bt == 0),
                    key=f"vis_3d_{gruppe}"
                )

        fig_3d = go.Figure()
        for gruppe, aktiv in aktive_gruppen.items():
            if not aktiv or gruppe not in meshes_3d:
                continue
            verts = meshes_3d[gruppe]["verts"]
            faces = meshes_3d[gruppe]["faces"]
            stil = BIM_FARBEN[gruppe]

            fig_3d.add_trace(go.Mesh3d(
                x=verts[:, 0],
                y=verts[:, 1],
                z=verts[:, 2],
                i=faces[:, 0],
                j=faces[:, 1],
                k=faces[:, 2],
                color=stil["farbe"],
                opacity=1.0,  # Komplett blickdicht (keine durchscheinenden Innenwände)
                name=f"{gruppe} ({meshes_3d[gruppe]['bauteile']})",
                showlegend=True,
                flatshading=True,
                lighting=dict(ambient=0.62, diffuse=0.88, roughness=0.45, specular=0.25),
                hovertemplate=f"<b>{gruppe}</b><br>X: %{{x:.2f}} m<br>Y: %{{y:.2f}} m<br>Z: %{{z:.2f}} m<extra></extra>",
            ))

        fig_3d.update_layout(
            height=620,
            margin=dict(l=0, r=0, t=10, b=0),
            paper_bgcolor="#E4EDF7",
            font=dict(color="#0F172A"),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.01,
                xanchor="left",
                x=0,
                bgcolor="rgba(255, 255, 255, 0.92)",
                bordercolor="#94A3B8",
                borderwidth=1,
                font=dict(color="#0F172A")
            ),
            scene=dict(
                aspectmode="data",
                xaxis=dict(
                    title="X (m)",
                    backgroundcolor="#D5E2F0",
                    gridcolor="#B4C6DC",
                    zerolinecolor="#0284C7",
                    color="#0F172A",
                    showbackground=True
                ),
                yaxis=dict(
                    title="Y (m)",
                    backgroundcolor="#D5E2F0",
                    gridcolor="#B4C6DC",
                    zerolinecolor="#0284C7",
                    color="#0F172A",
                    showbackground=True
                ),
                zaxis=dict(
                    title="Höhe Z (m)",
                    backgroundcolor="#C7D7EA",
                    gridcolor="#B4C6DC",
                    zerolinecolor="#0284C7",
                    color="#0F172A",
                    showbackground=True
                ),
                camera=dict(eye=dict(x=1.45, y=-1.45, z=0.85)),
            ),
        )

        st.plotly_chart(fig_3d, use_container_width=True)

        if modell_gekurzt:
            st.caption(
                f"Hinweis: Sehr detailliertes Modell – der 3D-Viewport wurde aus Performance-Gründen "
                f"auf {dreiecke_anzahl:,} Dreiecke begrenzt (die Mengenberechnung unten ist davon nicht betroffen)."
            )


# ==============================================================================
# 6. DIN 276 KOSTENGRUPPEN-ZUORDNUNG (2. EBENE) & EIGENE PREISLISTE
# ==============================================================================
st.markdown("---")
st.markdown("### DIN 276 Kostengruppen (2. Ebene) – Eigene Einheitspreise hinterlegen")
st.caption(
    "Die ermittelten IFC-Mengen sind nach DIN 276 (2. Ebene) gegliedert. "
    "Hinterlegen Sie Ihre eigenen Einheitspreise (z. B. aus BKI, Sirados oder Rahmenverträgen) direkt in der Tabelle "
    "oder laden Sie Ihre eigene Preisliste als CSV-Datei hoch:"
)

if not df_tueren.empty:
    anzahl_aussentueren = int(df_tueren["Türart"].isin(["Hauseingangstür", "Kelleraußentür"]).sum())
    anzahl_innentueren = int((~df_tueren["Türart"].isin(["Hauseingangstür", "Kelleraußentür"])).sum())
else:
    anzahl_aussentueren = 0
    anzahl_innentueren = 0

if not df_gelaender.empty:
    laenge_absturz_gelaender = float(
        df_gelaender.loc[df_gelaender["Geländerart"] == "Absturzgeländer (Fenster)", "Länge (m)"].sum()
    )
    laenge_treppen_gelaender = float(
        df_gelaender.loc[df_gelaender["Geländerart"] == "Treppengeländer", "Länge (m)"].sum()
    )
else:
    laenge_absturz_gelaender = 0.0
    laenge_treppen_gelaender = 0.0

if not df_decken.empty:
    flaeche_bodenplatte_kg320 = float(
        df_decken.loc[df_decken["Deckenart"] == "Bodenplatte / Kellerdecke", "Fläche (m²)"].sum()
    )
    flaeche_geschossdecken_kg350 = float(
        df_decken.loc[df_decken["Deckenart"] == "Geschossdecke", "Fläche (m²)"].sum()
    )
    flaeche_ogd_kg350 = float(
        df_decken.loc[df_decken["Deckenart"] == "Oberste Geschossdecke", "Fläche (m²)"].sum()
    )
else:
    flaeche_bodenplatte_kg320 = 0.0
    flaeche_geschossdecken_kg350 = 0.0
    flaeche_ogd_kg350 = 0.0

din276_positionen = [
    ("KG 320 – Gründung, Unterbau", "Sockeldämmung (Perimeter / Spritzwasser)", round(float(sockel.Fläche.sum()), 1), "m²"),
    ("KG 320 – Gründung, Unterbau", "Bodenplatte / Kellerdecke (Dämmung / Sanierung)", round(flaeche_bodenplatte_kg320, 1), "m²"),
    ("KG 330 – Außenwände / Vertikale Baukonstruktionen, außen", "Fassadendämmung (WDVS / Außenwand)", round(float(fassade.Fläche.sum()), 1), "m²"),
    ("KG 330 – Außenwände / Vertikale Baukonstruktionen, außen", "Fenster (Austausch inkl. Montage)", round(float(fenster_gesamt), 1), "m²"),
    ("KG 330 – Außenwände / Vertikale Baukonstruktionen, außen", "Außentüren (Haus- & Kelleraußentüren)", float(anzahl_aussentueren), "Stk."),
    ("KG 330 – Außenwände / Vertikale Baukonstruktionen, außen", "Absturzgeländer (Fenster / Balkon)", round(laenge_absturz_gelaender, 1), "m"),
    ("KG 340 – Innenwände / Vertikale Baukonstruktionen, innen", "Innentüren (Wohnung / Keller / DG)", float(anzahl_innentueren), "Stk."),
    ("KG 340 – Innenwände / Vertikale Baukonstruktionen, innen", "Treppengeländer / Handläufe", round(laenge_treppen_gelaender, 1), "m"),
    ("KG 350 – Decken / Horizontale Baukonstruktionen", "Geschossdecken (Belag / Trittschall / Sanierung)", round(flaeche_geschossdecken_kg350, 1), "m²"),
    ("KG 350 – Decken / Horizontale Baukonstruktionen", "Oberste Geschossdecke (Wärmedämmung)", round(flaeche_ogd_kg350, 1), "m²"),
    ("KG 360 – Dächer", "Dachfläche (Dämmung & Neueindeckung)", round(float(dach_gesamt), 1), "m²"),
    ("KG 360 – Dächer", "Regenrinnen", round(float(rinnen_gesamt), 1), "m"),
    ("KG 360 – Dächer", "Fallrohre", round(float(fallrohre_gesamt), 1), "m"),
]

importierte_preise = {}
importierte_quellen = {}

with st.expander("Eigene Preisliste als CSV hochladen oder Blanko-Vorlage herunterladen", expanded=False):
    col_csv1, col_csv2 = st.columns(2)
    with col_csv1:
        vorlage_df = pd.DataFrame([
            {
                "Kostengruppe": kg,
                "Bauteil / Maßnahme": bauteil,
                "Einheit": einheit,
                "Einheitspreis_Netto": 0.0,
                "Preisquelle_Stand": "z. B. BKI 2026 / Rahmenvertrag",
            }
            for kg, bauteil, _, einheit in din276_positionen
        ])
        csv_vorlage_bytes = vorlage_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
        st.download_button(
            "Blanko-Preisliste herunterladen (.csv)",
            data=csv_vorlage_bytes,
            file_name="din276_preisvorlage.csv",
            mime="text/csv",
            help="Speichern Sie diese Datei auf Ihrem Rechner, tragen Sie Ihre Einheitspreise in Excel ein und laden Sie sie rechts wieder hoch."
        )
    with col_csv2:
        preis_csv = st.file_uploader(
            "Ausgefüllte Preisliste hochladen (.csv)",
            type=["csv"],
            key="upload_din276_csv"
        )
        if preis_csv is not None:
            try:
                df_csv = pd.read_csv(preis_csv, sep=None, engine="python", decimal=",")
                if {"Bauteil / Maßnahme", "Einheitspreis_Netto"}.issubset(df_csv.columns):
                    for _, row in df_csv.iterrows():
                        bt_key = str(row["Bauteil / Maßnahme"]).strip()
                        preis_val = pd.to_numeric(str(row["Einheitspreis_Netto"]).replace(",", "."), errors="coerce")
                        if pd.notna(preis_val) and preis_val >= 0:
                            importierte_preise[bt_key] = float(preis_val)
                        if "Preisquelle_Stand" in df_csv.columns and pd.notna(row["Preisquelle_Stand"]):
                            importierte_quellen[bt_key] = str(row["Preisquelle_Stand"]).strip()
                    st.success(f"{len(importierte_preise)} Einheitspreise aus CSV übernommen.")
                else:
                    st.error("Die CSV-Datei benötigt mindestens die Spalten „Bauteil / Maßnahme“ und „Einheitspreis_Netto“.")
            except Exception as exc:
                st.error(f"Fehler beim Lesen der Preisliste: {exc}")

df_din276_basis = pd.DataFrame([
    {
        "Kostengruppe (DIN 276)": kg,
        "Bauteil / Maßnahme": bauteil,
        "Menge": menge,
        "Einheit": einheit,
        "Einheitspreis (€ netto)": importierte_preise.get(bauteil, 0.0),
        "Preisquelle / Stand": importierte_quellen.get(bauteil, ""),
    }
    for kg, bauteil, menge, einheit in din276_positionen
])

df_din276_edit = st.data_editor(
    df_din276_basis,
    use_container_width=True,
    hide_index=True,
    disabled=["Kostengruppe (DIN 276)", "Bauteil / Maßnahme", "Menge", "Einheit"],
    column_config={
        "Menge": st.column_config.NumberColumn(format="%.1f"),
        "Einheitspreis (€ netto)": st.column_config.NumberColumn(
            min_value=0.0,
            step=5.0,
            format="%.2f €",
            help="Klicken Sie in eine Zelle, um Ihren eigenen Einheitspreis (netto) einzutragen."
        ),
        "Preisquelle / Stand": st.column_config.TextColumn(
            help="Optional: Notieren Sie hier Ihre Quelle (z. B. BKI 2026, Angebot Firma X), damit sie im Excel-Export dokumentiert ist."
        ),
    },
    key=f"editor_din276_{preis_csv.name if preis_csv else 'manuell'}"
)

df_din276_edit["Gesamtkosten (€ netto)"] = (
    df_din276_edit["Menge"] * df_din276_edit["Einheitspreis (€ netto)"]
).round(2)

zus_din276_kg = (
    df_din276_edit.groupby("Kostengruppe (DIN 276)", as_index=False)["Gesamtkosten (€ netto)"]
    .sum()
)
gesamtsumme_netto = float(zus_din276_kg["Gesamtkosten (€ netto)"].sum())
zus_din276_kg.loc[len(zus_din276_kg)] = ["Summe Baukonstruktionen (KG 300 netto)", round(gesamtsumme_netto, 2)]

c_kg1, c_kg2 = st.columns([2, 1])
with c_kg1:
    st.markdown("**Zusammenfassung nach Kostengruppen (2. Ebene)**")
    st.table(
        zus_din276_kg.style.hide(axis="index").format({"Gesamtkosten (€ netto)": "{:,.2f} €"})
    )
with c_kg2:
    st.markdown(
        f'<div class="karte">'
        f'<span>Summe hinterlegte Preise (KG 300 netto)</span>'
        f'<b>{gesamtsumme_netto:,.2f} €</b>'
        f'</div>',
        unsafe_allow_html=True
    )
    st.caption(
        "Alle Preise basieren ausschließlich auf Ihren eigenen Eingaben (netto, zzgl. Baunebenkosten KG 700 und MwSt.). "
        "Ihre eingetragenen Werte verbleiben ausschließlich in Ihrer Sitzung und werden nicht auf dem Server gespeichert."
    )


# ==============================================================================
# 7. DETAILTABELLEN & EXCEL-EXPORT
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

if not df_decken.empty:
    zus_decken = df_decken.groupby(["Geschoss", "Deckenart"]).agg(
        Anzahl=("Name", "size"),
        Fläche=("Fläche (m²)", "sum")
    ).reset_index().rename(columns={"Fläche": "Fläche (m²)"})
    zus_decken.loc[len(zus_decken)] = ["Gesamt", "Alle Decken", zus_decken["Anzahl"].sum(), zus_decken["Fläche (m²)"].sum()]
else:
    zus_decken = pd.DataFrame(columns=["Geschoss", "Deckenart", "Anzahl", "Fläche (m²)"])

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
detail_decken = (
    df_decken[["Name", "GlobalId", "Geschoss", "Deckenart", "Fläche (m²)", "Herkunft"]]
    if not df_decken.empty
    else pd.DataFrame(columns=["Name", "GlobalId", "Geschoss", "Deckenart", "Fläche (m²)", "Herkunft"])
)
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

st.markdown("---")
st.markdown("### Wände: Zusammenfassung")
st.table(fmt(zus_waende))
st.markdown("### Wände: Pro Geschoss")
st.table(fmt(geo_waende))

st.markdown("### Decken & Bodenplatten: Zusammenfassung")
if not zus_decken.empty:
    st.table(fmt(zus_decken))
else:
    st.info("Keine Geschossdecken oder Bodenplatten (IfcSlab) im Modell gefunden.")

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
with st.expander("Alle Decken & Bodenplatten einzeln"):
    st.dataframe(detail_decken, width="stretch", hide_index=True)
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
        "Decken & Bodenplatten gesamt (m²)",
        "Fensterfläche gesamt (m²)",
        "Türen gesamt (Anzahl)",
        "Dachfläche gesamt (m²)",
        "Geländer gesamt (m)",
        "Regenrinnen gesamt (m)",
        "Fallrohre gesamt (m)",
        "Kostenermittlung KG 300 netto (€)"
    ],
    "Wert": [
        round(fassade.Fläche.sum(), 1),
        round(sockel.Fläche.sum(), 1),
        round(decken_gesamt, 1),
        round(fenster_gesamt, 1),
        int(tueren_anzahl),
        round(dach_gesamt, 1),
        round(gelaender_gesamt, 1),
        round(rinnen_gesamt, 1),
        round(fallrohre_gesamt, 1),
        round(gesamtsumme_netto, 2)
    ]
})

buf = io.BytesIO()
sheets = {
    "Ergebnis": ergebnis,
    "DIN 276 Kostengruppen": df_din276_edit,
    "DIN 276 Summen": zus_din276_kg,
    "Wände Zusammenfassung": zus_waende,
    "Wände pro Geschoss": geo_waende,
    "Decken Zusammenfassung": zus_decken,
    "Fenster Zusammenfassung": zus_fenster,
    "Türen Zusammenfassung": zus_tueren,
    "Dächer Zusammenfassung": zus_daecher,
    "Geländer Zusammenfassung": zus_gelaender,
    "Rinnen & Fallrohre": zus_entwaesserung,
    "Alle Wände": detail_waende,
    "Alle Decken": detail_decken,
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
            c.fill = PatternFill("solid", fgColor=EXCEL_HEADER)
            c.font = Font(color="FFFFFF", bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(60, max(len(str(c.value or "")) for c in col) + 3)

col_dl, col_clear = st.columns([2, 1])
with col_dl:
    st.download_button(
        "Als Excel exportieren",
        buf.getvalue(),
        "bauteilauswertung-ifc.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with col_clear:
    if st.button("Sitzung & Arbeitsspeicher leeren", use_container_width=True, help="Entfernt alle geladenen Modelldaten sofort aus dem Arbeitsspeicher."):
        st.cache_data.clear()
        st.rerun()
