# ======================================================================
# Portail Génie Civil — EDT & Listes (S1)
# Fichier : app/streamlit_app.py
# ======================================================================
# - Profils Étudiant / Enseignant
# - Filtres Spécialité → Niveau → Groupe
# - Normalisation EDT & listes étudiants (S1)
# - Inférence robuste depuis nom de fichier (2ING, ING2, etc.)
# - Export Excel (.xlsx) pour EDT/Planning
# - Feuille de présence mobile (sans matricule) + Remarque + Tout cocher/décocher
# - Export PDF présence avec en-tête institutionnel
# - Panneaux diagnostic masqués par défaut
# ======================================================================

from __future__ import annotations

import re
import glob
from io import BytesIO
from datetime import datetime, timedelta, time as dtime
from pathlib import Path
from typing import Tuple, Optional, Iterable, Dict, Any

import pandas as pd
import streamlit as st

# --- pour l'export PDF ---
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors

# --------------------------- CONFIG GLOBALE ----------------------------

st.set_page_config(
    page_title="EDT & Listes • Génie Civil (S1)",
    page_icon="🗓️",
    layout="wide",
)

BASE_DIR   = Path(__file__).resolve().parent
RAW_EDT    = str(BASE_DIR / "data" / "raw" / "edt")
RAW_STU    = str(BASE_DIR / "data" / "raw" / "students")
SEMESTRE   = "S1"                 # application mono-semestre
SHOW_DIAGNOSTIC = False           # <- mettre True pour réafficher l’expander diagnostic

# Ordres & libellés
ORDER_JOUR = {
    "DIMANCHE": 0, "LUNDI": 1, "MARDI": 2, "MERCREDI": 3,
    "JEUDI": 4, "VENDREDI": 5, "SAMEDI": 6,
}
JOURS_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]

# Colonnes attendues
EDT_COLS = [
    "Niveau", "Spécialité", "Groupe", "Semestre", "Jour", "Heure début", "Heure fin",
    "Durée (h)", "Matière", "Type", "Enseignant", "Salle", "Fréquence",
]
STU_COLS = [
    "Annee", "Semestre", "Spécialité", "Niveau", "Groupe",
    "Matricule", "Nom", "Prenom", "Email", "Téléphone", "Remarque", "N°",
]

# ------------------------------ STYLES --------------------------------

def inject_css() -> None:
    st.markdown(
        """
        <style>
          /* Commun */
          h1, h2, h3 { letter-spacing:.2px }
          .actionbar { display:flex; gap:.5rem; flex-wrap:wrap; margin:.25rem 0 1rem }
          .pill { padding:.38rem .65rem; border-radius:999px; font-size:.86rem; font-weight:600; border:1px solid transparent; }
          .badge { padding:.12rem .45rem; border-radius:8px; font-size:.78rem; margin-left:.35rem; border:1px solid transparent; font-weight:600; }
          .card  { border:1px solid transparent; padding:1rem; border-radius:12px; margin:.25rem 0 .75rem; }
          .muted{ opacity:.85; font-size:.92rem }
          .next-title { font-size:1.1rem; font-weight:700; }

          /* Sombre */
          @media (prefers-color-scheme: dark) {
            .pill { background:#1f2937; color:#e5e7eb; border-color:#374151; }
            .role-etudiant  { background:#0e3c2e; color:#a7f3d0; border-color:#065f46; }
            .role-enseignant{ background:#2b2547; color:#c7d2fe; border-color:#4338ca; }

            .card { background:#0f1624; border-color:#1f2937; }
            .badge { background:#1f2937; color:#e5e7eb; border-color:#374151; }

            .next-title { color:#e5e7eb; }
          }

          /* Clair */
          @media (prefers-color-scheme: light) {
            .pill { background:#eef2ff; color:#1f2937; border-color:#c7d2fe; }
            .role-etudiant  { background:#dcfce7; color:#064e3b; border-color:#86efac; }
            .role-enseignant{ background:#e0e7ff; color:#312e81; border-color:#a5b4fc; }

            .card { background:#f8fafc; border-color:#e5e7eb; }
            .badge { background:#f1f5f9; color:#111827; border-color:#e5e7eb; }

            .next-title { color:#111827; }
          }

          /* Data Editor (présence) */
          [data-testid="stDataEditorContainer"] { border-radius: 12px; }
          [data-testid="stDataEditorRow"] { min-height: 40px; }
          [data-testid="column-Nom complet"] div { white-space: normal !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

def header_role(role_label: str, subtitle: str) -> None:
    role_class = "role-etudiant" if role_label == "Étudiant" else "role-enseignant"
    st.markdown(
        f"""
        <div class="actionbar">
          <span class="pill {role_class}">
            {"👩‍🎓 Étudiant" if role_label == "Étudiant" else "👨‍🏫 Enseignant"}
          </span>
          <span class="pill">{subtitle}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

inject_css()

# ----------------------------- HELPERS --------------------------------

def read_any(path: str) -> pd.DataFrame:
    """Lecture de fichiers .xlsx ou .csv."""
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path)

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame → bytes .xlsx (openpyxl)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
    return buf.getvalue()

def time_to_minutes(h: Any) -> Optional[int]:
    """'08h30' → 510. Retourne None si invalide."""
    s = str(h).strip().lower().replace(" ", "")
    if "h" not in s:
        return None
    try:
        hh, mm = s.split("h")
        return int(hh or 0) * 60 + int(mm or 0)
    except Exception:
        return None

def minutes_to_dt(d: datetime, minutes: int) -> datetime:
    return datetime.combine(d.date(), dtime.min) + timedelta(minutes=minutes)

def human_delta(dt: datetime, now: datetime) -> str:
    s = int((dt - now).total_seconds())
    d = s // 86400; s %= 86400
    h = s // 3600; s %= 3600
    m = s // 60
    out = []
    if d: out.append(f"{d}j")
    if h: out.append(f"{h}h")
    if m: out.append(f"{m}m")
    return " ".join(out) or "0m"
# --------------------- NORMALISATION & INFERENCE ----------------------

def ensure_cols(df: pd.DataFrame, cols: Iterable[str], numeric: Iterable[str] = ()) -> pd.DataFrame:
    numeric = set(numeric or [])
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0 if c in numeric else ""
    return df[[c for c in cols]]

def normalize_semestre(val: Any, fallback: str = "S1") -> str:
    v = str(val).strip().upper()
    return v or fallback

def normalize_groupe(val: Any) -> str:
    s = str(val).upper().replace(" ", "")
    if not s:
        return s
    if not s.startswith("G") and s.isdigit():
        return "G" + s
    return s

def classify_spec_level(spec_text: str, level_text: str) -> Tuple[str, str]:
    S = (spec_text or "").upper()
    L = (level_text or "").upper()
    if "RIB" in S:        return "RIB",        "M1" if "M1" in S+L else ("M2" if "M2" in S+L else "")
    if "VOA" in S:        return "VOA",        "M1" if "M1" in S+L else ("M2" if "M2" in S+L else "")
    if "STRUCT" in S:     return "STRUCTURE",  "M1" if "M1" in S+L else ("M2" if "M2" in S+L else "")
    if "L2" in S+L:       return "LICENCE", "2"
    if "L3" in S+L:       return "LICENCE", "3"
    if any(k in S+L for k in ["ING", "INGÉ", "INGENIEUR", "INGÉNIEUR"]):
        if "1" in S+L: return "INGENIEUR", "1"
        if "2" in S+L: return "INGENIEUR", "2"
        if "3" in S+L: return "INGENIEUR", "3"
        return "INGENIEUR", ""
    return "", ""

# ----------------- INFÉRENCE MÉTA DEPUIS LE NOM DU FICHIER ------------

def infer_from_filename(path: str) -> Tuple[Optional[str], Optional[str], Optional[str], str]:
    name = Path(path).stem.upper()
    name_compact = re.sub(r"[\s\-]+", "_", name)
    name_nospace = re.sub(r"\s+", "", name)

    # Groupe
    g = None
    m = re.search(r"(?:^|[_\-])G\s*?(\d+)", name, flags=re.I)
    if not m:
        m = re.search(r"G\s*?(\d+)", name, flags=re.I)
    if m:
        g = f"G{m.group(1)}"

    # Semestre
    sem = None
    m = re.search(r"(?:^|[_\-])S\s*?(\d+)", name, flags=re.I)
    if not m:
        m = re.search(r"S\s*?(\d+)", name, flags=re.I)
    if m:
        sem = f"S{m.group(1)}"

    # Spécialités M1/M2
    if "RIB" in name_nospace:
        niv = "M2" if "M2" in name_nospace else ("M1" if "M1" in name_nospace else "")
        return "RIB", niv, g, (sem or "S1")
    if "VOA" in name_nospace:
        niv = "M2" if "M2" in name_nospace else ("M1" if "M1" in name_nospace else "")
        return "VOA", niv, g, (sem or "S1")
    if "STRUCT" in name_nospace or "STRUC" in name_nospace:
        niv = "M2" if "M2" in name_nospace else ("M1" if "M1" in name_nospace else "")
        return "STRUCTURE", niv, g, (sem or "S1")

    # LICENCE 2/3
    m = re.search(r"(?:LICENCE|L)[_\- ]?([23])", name_compact)
    if m:
        return "LICENCE", m.group(1), g, (sem or "S1")

    # INGENIEUR 1/2/3
    m = re.search(r"(?:INGENIEUR|ING)[_\- ]?([123])", name_compact)
    if not m:
        m = re.search(r"([123])[_\- ]?(?:INGENIEUR|ING)", name_compact)
    if m:
        return "INGENIEUR", m.group(1), g, (sem or "S1")

    return None, None, g, (sem or "S1")

# ------------------ HARMONISATION LISTES ÉTUDIANTS --------------------

def harmonize_student_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping: Dict[str, str] = {}
    for c in df.columns:
        k = str(c).strip().lower()
        if k in {"n°", "nº", "n", "num", "numero", "numéro", "n°/ordre"}:
            mapping[c] = "N°"
        elif k in {"matricule", "code", "id", "cne", "apogee", "apogée"}:
            mapping[c] = "Matricule"
        elif k == "nom":
            mapping[c] = "Nom"
        elif k in {"prenom", "prénom"}:
            mapping[c] = "Prenom"
        elif k in {"nom et prénom", "nom et prenom", "nom_prenom", "nom-prenom"}:
            mapping[c] = "NomPrenom"
        elif k.startswith("remarq") or k.startswith("obs"):
            mapping[c] = "Remarque"
        elif k.startswith("semestre"):
            mapping[c] = "Semestre"
        elif k.startswith("groupe"):
            mapping[c] = "Groupe"
        elif k.startswith("spécial") or k.startswith("special"):
            mapping[c] = "Spécialité"
        elif k in {"annee","année","annee scolaire","année scolaire"}:
            mapping[c] = "Annee"

    if mapping:
        df = df.rename(columns=mapping)

    if "NomPrenom" in df.columns and (("Nom" not in df.columns) or ("Prenom" not in df.columns)):
        np = df["NomPrenom"].astype(str).str.strip()
        if "Nom" not in df.columns:
            df["Nom"] = np.str.split(r"\s+", n=1, expand=True)[0].fillna("")
        if "Prenom" not in df.columns:
            part = np.str.split(r"\s+", n=1, expand=True)
            df["Prenom"] = (part[1] if part.shape[1] > 1 else "").fillna("")

    for col in ["N°", "Matricule", "Nom", "Prenom", "Remarque", "Semestre", "Spécialité", "Niveau", "Groupe"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ------------------ CHARGEMENT ET MISE EN FORME (PATCHÉ) --------------

@st.cache_data
def load_raw_s1() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    idx_rows = []

    # ----------------------------- EDT ---------------------------------
    edt_frames = []
    for f in sorted(glob.glob(f"{RAW_EDT}/*.*")):
        fname = Path(f).name
        name_up = fname.upper()

        try:
            df = read_any(f)
        except Exception as e:
            idx_rows.append({"Type":"EDT","Fichier":fname,"Lu":False,"Erreur":str(e),
                             "S1_par_nom":"S1" in name_up,"S1_par_col":None,
                             "Spec2":None,"Niv2":None,"Groupe":None})
            continue

        df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
        df["Semestre"] = df["Semestre"].apply(normalize_semestre)
        df["Groupe"]   = df["Groupe"].apply(normalize_groupe)

        has_s1_col = df["Semestre"].astype(str).str.upper().eq("S1").any()
        s1_by_name = "S1" in name_up
        if not (s1_by_name or has_s1_col):
            idx_rows.append({"Type":"EDT","Fichier":fname,"Lu":True,"Erreur":"",
                             "S1_par_nom":s1_by_name,"S1_par_col":has_s1_col,
                             "Spec2":None,"Niv2":None,"Groupe":None})
            continue

        specs, nivs = zip(*df.apply(lambda r: classify_spec_level(r.get("Spécialité",""),
                                                                  r.get("Niveau","")), axis=1))
        df["Spec2"], df["Niv2"] = specs, nivs

        # --- INFÉRENCE + FORÇAGE GLOBAL (PATCH)
        s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
        df["Spec2"]    = df["Spec2"].fillna("").astype(str).str.upper()
        df["Niv2"]     = df["Niv2"].fillna("").astype(str).str.upper()
        df["Groupe"]   = df["Groupe"].fillna("").astype(str)
        df["Semestre"] = df["Semestre"].fillna("").astype(str).str.upper()
        if s2_f:  df["Spec2"]    = s2_f
        if n2_f:  df["Niv2"]     = n2_f
        if g_f:   df["Groupe"]   = normalize_groupe(g_f)
        if sem_f: df["Semestre"] = sem_f

        idx_rows.append({
            "Type":"EDT","Fichier":fname,"Lu":True,"Erreur":"",
            "S1_par_nom":s1_by_name,"S1_par_col":has_s1_col,
            "Spec2":df["Spec2"].dropna().astype(str).str.upper().replace("",None).mode().tolist()[:1] or [None],
            "Niv2": df["Niv2"].dropna().astype(str).str.upper().replace("",None).mode().tolist()[:1] or [None],
            "Groupe":df["Groupe"].dropna().map(normalize_groupe).mode().tolist()[:1] or [None],
        })

        df = df[df["Semestre"].astype(str).str.upper() == "S1"].copy()
        edt_frames.append(df)

    if edt_frames:
        edt = pd.concat(edt_frames, ignore_index=True)
        edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
        edt = edt.sort_values(["Spec2","Niv2","Groupe","__o","Heure début"]).drop(columns="__o")
    else:
        edt = pd.DataFrame(columns=EDT_COLS + ["Spec2","Niv2"])

    # --------------------------- ETUDIANTS -------------------------------
    stu_frames = []
    for f in sorted(glob.glob(f"{RAW_STU}/*.*")):
        fname = Path(f).name
        name_up = fname.upper()

        try:
            df = read_any(f)
        except Exception as e:
            idx_rows.append({"Type":"ETU","Fichier":fname,"Lu":False,"Erreur":str(e),
                             "S1_par_nom":"S1" in name_up,"S1_par_col":None,
                             "Spec2":None,"Niv2":None,"Groupe":None})
            continue

        df = harmonize_student_columns(df)
        df = ensure_cols(df, STU_COLS)
        df["Semestre"] = df["Semestre"].apply(normalize_semestre)
        df["Groupe"]   = df["Groupe"].apply(normalize_groupe)

        has_s1_col = df["Semestre"].astype(str).str.upper().eq("S1").any()
        s1_by_name = "S1" in name_up
        if not (s1_by_name or has_s1_col):
            idx_rows.append({"Type":"ETU","Fichier":fname,"Lu":True,"Erreur":"",
                             "S1_par_nom":s1_by_name,"S1_par_col":has_s1_col,
                             "Spec2":None,"Niv2":None,"Groupe":None})
            continue

        specs, nivs = zip(*df.apply(lambda r: classify_spec_level(r.get("Spécialité",""),
                                                                  r.get("Niveau","")), axis=1))
        df["Spec2"], df["Niv2"] = specs, nivs

        # --- INFÉRENCE + FORÇAGE GLOBAL (PATCH)
        s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
        df["Spec2"]    = df["Spec2"].fillna("").astype(str).str.upper()
        df["Niv2"]     = df["Niv2"].fillna("").astype(str).str.upper()
        df["Groupe"]   = df["Groupe"].fillna("").astype(str)
        df["Semestre"] = df["Semestre"].fillna("").astype(str).str.upper()
        if s2_f:  df["Spec2"]    = s2_f
        if n2_f:  df["Niv2"]     = n2_f
        if g_f:   df["Groupe"]   = normalize_groupe(g_f)
        if sem_f: df["Semestre"] = sem_f

        idx_rows.append({
            "Type":"ETU","Fichier":fname,"Lu":True,"Erreur":"",
            "S1_par_nom":s1_by_name,"S1_par_col":has_s1_col,
            "Spec2":df["Spec2"].dropna().astype(str).str.upper().replace("",None).mode().tolist()[:1] or [None],
            "Niv2": df["Niv2"].dropna().astype(str).str.upper().replace("",None).mode().tolist()[:1] or [None],
            "Groupe":df["Groupe"].dropna().map(normalize_groupe).mode().tolist()[:1] or [None],
        })

        df = df[df["Semestre"].astype(str).str.upper() == "S1"].copy()
        stu_frames.append(df)

    etu = pd.concat(stu_frames, ignore_index=True) if stu_frames else pd.DataFrame(columns=STU_COLS + ["Spec2","Niv2"])

    # --------------------- TABLE D’INDEX (affichage) -------------------
    idx = pd.DataFrame(idx_rows)
    for c in ["Spec2","Niv2","Groupe"]:
        if c in idx.columns:
            idx[c] = idx[c].apply(lambda v: v[0] if isinstance(v, list) and v else v)

    return edt, etu, idx


def subgroup_by_spec_level(df: pd.DataFrame, spec: Optional[str]=None,
                           niv: Optional[str]=None, groupe: Optional[str]=None) -> pd.DataFrame:
    keep = df[df["Semestre"].astype(str).str.upper() == SEMESTRE]
    if spec:
        keep = keep[keep["Spec2"] == spec]
    if niv:
        keep = keep[keep["Niv2"] == niv]
    if groupe:
        gnorm = normalize_groupe(groupe)
        keep = keep[keep["Groupe"].apply(normalize_groupe) == gnorm]
    return keep


def level_options_for(spec: str) -> Iterable[str]:
    if spec in ("RIB","VOA","STRUCTURE"): return ["M1","M2"]
    if spec == "LICENCE": return ["2","3"]
    if spec == "INGENIEUR": return ["1","2","3"]
    return []

def pretty_level_label(spec: str, niv: str) -> str:
    if spec == "LICENCE": return f"LICENCE {niv}"
    if spec == "INGENIEUR": return f"INGENIEUR {niv}"
    return niv

# -------------------------- PDF DE PRÉSENCE ---------------------------

def make_presence_pdf(
    df: pd.DataFrame,
    titre: str,
    meta: str,
    header: Optional[dict] = None,
) -> bytes:
    """
    Génère un PDF A4 :
    - entête institutionnelle (Université / Faculté / Département / Spécialité + Groupe)
    - titre + méta (date/heure)
    - tableau : N°, Nom complet, Présent, Remarque
    """
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    margin = 1.2 * cm
    y = height - margin

    # -------- Entête institutionnelle (centrée) --------
    if header is None:
        header = {}
    univ = header.get("univ", "UNIVERSITÉ DE TLEMCEN").upper()
    fac  = header.get("fac",  "FACULTÉ DE TECHNOLOGIE").upper()
    dept = header.get("dept", "DÉPARTEMENT DE GÉNIE CIVIL").upper()
    spec_line = header.get("spec", "").strip()       # ex: "Spécialité : RIB — Niveau : M1"
    grp_line  = header.get("grp", "").strip()        # ex: "Groupe : G11"

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, univ); y -= 14
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(width/2, y, fac);  y -= 13
    c.drawCentredString(width/2, y, dept); y -= 15

    if spec_line or grp_line:
        c.setFont("Helvetica", 10)
        if spec_line:
            c.drawCentredString(width/2, y, spec_line); y -= 12
        if grp_line:
            c.drawCentredString(width/2, y, grp_line);  y -= 12

    # petite ligne de séparation
    c.setStrokeColor(colors.grey)
    c.line(margin, y, width - margin, y)
    y -= 10
    c.setStrokeColor(colors.black)

    # -------- Titre + méta (gauche) --------
    c.setFont("Helvetica-Bold", 13)
    c.drawString(margin, y, titre)
    y -= 16
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.black)
    c.drawString(margin, y, meta)
    y -= 18

    # Dimensions colonnes : N° | Nom complet | Présent | Remarque
    w_num = 1.0 * cm
    w_pres = 2.0 * cm
    w_nom = width - 2 * margin - w_num - w_pres - 6.2 * cm
    if w_nom < 6.0 * cm:
        w_nom = 6.0 * cm
    w_rem = width - 2 * margin - w_num - w_pres - w_nom

    row_h = 13

    def new_page_header(page_idx: int):
        """Entête de page : titre + méta (pages suivantes)."""
        nonlocal y
        if y < margin + 4 * row_h:
            c.showPage()
            y = height - margin
            c.setFont("Helvetica-Bold", 13)
            c.drawString(margin, y, f"{titre}  (p.{page_idx})")
            y -= 16
            c.setFont("Helvetica", 10)
            c.drawString(margin, y, meta)
            y -= 18

        # entête tableau
        c.setFont("Helvetica-Bold", 10)
        x = margin
        for header_txt, w in [("N°", w_num), ("Nom complet", w_nom), ("Présent", w_pres), ("Remarque", w_rem)]:
            c.rect(x, y - row_h, w, row_h, stroke=1, fill=0)
            c.drawString(x + 3, y - row_h + 3, header_txt)
            x += w
        y -= row_h
        c.setFont("Helvetica", 9)

    page_idx = 1
    new_page_header(page_idx)

    # Corps du tableau
    for i, r in df.reset_index(drop=True).iterrows():
        if y < margin + row_h:
            page_idx += 1
            new_page_header(page_idx)
        x = margin
        vals = [
            str(i + 1),
            str(r.get("Nom complet", "")),
            ("✔" if bool(r.get("Présent", False)) else ""),
            str(r.get("Remarque", "") or ""),
        ]
        for (val, w) in zip(vals, [w_num, w_nom, w_pres, w_rem]):
            c.rect(x, y - row_h, w, row_h, stroke=1, fill=0)
            c.drawString(x + 3, y - row_h + 3, val[:120])
            x += w
        y -= row_h

    c.showPage()
    c.save()
    pdf = buf.getvalue()
    buf.close()
    return pdf
# ============================ INTERFACE ===============================

st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu, idx = load_raw_s1()

# ---- Index de détection (diagnostic complet) — masqué par défaut
if SHOW_DIAGNOSTIC:
    with st.expander("📂 Index des fichiers détectés (diagnostic)", expanded=False):
        if idx.empty:
            st.info("Aucun fichier détecté dans `data/raw/edt/` et `data/raw/students/`.")
        else:
            st.dataframe(
                idx[["Type","Fichier","Lu","S1_par_nom","S1_par_col","Spec2","Niv2","Groupe","Erreur"]],
                hide_index=True, use_container_width=True
            )

if edt.empty:
    st.error("Aucun EDT S1 trouvé ou reconnu.")
    st.stop()

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)

    st.markdown("---")
    st.caption("Filtres hiérarchiques")

    spec_order = ["RIB","VOA","STRUCTURE","LICENCE","INGENIEUR"]
    available_specs = [s for s in spec_order if s in edt["Spec2"].dropna().unique().tolist() or s in etu["Spec2"].dropna().unique().tolist()]
    spec = st.selectbox("Spécialité", available_specs, index=0 if available_specs else None)

    raw_levels = list(level_options_for(spec))
    level_labels = [pretty_level_label(spec, n) for n in raw_levels]
    label_to_raw = dict(zip(level_labels, raw_levels))
    niv_label = st.selectbox("Niveau", level_labels, index=0 if level_labels else None)
    niv = label_to_raw.get(niv_label)

    g_from_edt = subgroup_by_spec_level(edt, spec, niv)["Groupe"].dropna().map(normalize_groupe)
    g_from_etu = subgroup_by_spec_level(etu, spec, niv)["Groupe"].dropna().map(normalize_groupe)
    grp_pool = sorted(pd.concat([g_from_edt, g_from_etu]).unique().tolist())
    groupe = st.selectbox("Groupe", grp_pool, index=0 if grp_pool else None)

    st.markdown("---")
    q_nom = st.text_input("Nom/Prénom (étudiant ou enseignant)")
    st.caption("Astuce : tape un nom puis appuie sur Entrée ⏎")
    print_mode = st.checkbox("🖨️ Mode impression")

if print_mode:
    st.markdown(
        """
        <style>
            section[data-testid="stSidebar"] { display: none !important; }
            .block-container { padding-top: 1rem; padding-bottom: 0; }
            header { visibility: hidden; height: 0; }
            @media print {
                .stButton, .stDownloadButton, [data-testid="stFileUploader"] { display: none !important; }
                .stDataFrame { border: none !important; }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

bloc = subgroup_by_spec_level(edt, spec, niv, groupe)

now = datetime.now()
title_clean = f"{spec} {pretty_level_label(spec, niv)}".strip()
# ============================ VUE ÉTUDIANT ============================

if role == "Étudiant":
    header_role("Étudiant", f"{title_clean} • Groupe {groupe}")

    tab_edt, tab_next = st.tabs(["📅 Mon EDT", "⏭️ Prochaine séance"])

    with tab_edt:
        st.markdown("#### Emploi du temps")
        view = bloc[["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence"]]
        st.download_button(
            "⬇️ Exporter l’EDT en Excel",
            df_to_xlsx_bytes(view),
            file_name=f"EDT_{spec}_{niv}_G{groupe}_S1.xlsx",
            use_container_width=True,
        )
        st.dataframe(view.rename(columns={"Heure début":"Début","Heure fin":"Fin"}),
                     use_container_width=True, hide_index=True)

    with tab_next:
        st.markdown("#### À venir")

        def next_session(df: pd.DataFrame) -> Optional[Tuple[datetime, pd.Series]]:
            if df.empty:
                return None
            py_day = {"LUNDI":0,"MARDI":1,"MERCREDI":2,"JEUDI":3,"VENDREDI":4,"SAMEDI":5,"DIMANCHE":6}
            today_idx = now.weekday()
            rows = []
            for _, r in df.iterrows():
                d_idx = py_day.get(str(r["Jour"]).upper(), None)
                if d_idx is None: continue
                m = time_to_minutes(r["Heure début"])
                if m is None: continue
                delta = (d_idx - today_idx) % 7
                dt = datetime.combine((now + timedelta(days=delta)).date(), dtime.min) + timedelta(minutes=m)
                if dt < now:
                    dt += timedelta(days=7)
                rows.append((dt, r))
            if not rows: return None
            rows.sort(key=lambda x: x[0])
            return rows[0]

        nxt = next_session(bloc)
        if nxt:
            dt, r = nxt
            st.markdown(
                f"""
                <div class="card">
                  <div class="next-title">{r['Matière']} <span class="badge">{r['Type']}</span></div>
                  <div class="muted">
                    {r['Jour']} • {r['Heure début']}–{r['Heure fin']}
                    <span class="badge">Salle {r['Salle']}</span>
                    <span class="badge">Avec {r['Enseignant']}</span>
                    <span class="badge">Dans {human_delta(dt, now)}</span>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("Aucune séance à venir avec ces filtres.")

# =========================== VUE ENSEIGNANT ===========================

else:
    header_role("Enseignant", f"{title_clean} • Groupe {groupe}")

    tab_plan, tab_next, tab_where, tab_presence = st.tabs(
        ["🗂️ Planning", "⏭️ Prochaine séance", "📍 Où trouver un enseignant ?", "📝 Feuille de présence"]
    )

    with tab_plan:
        st.markdown("#### Planning filtré")
        planning = bloc.copy()
        if q_nom:
            planning = planning[planning["Enseignant"].str.contains(q_nom, case=False, na=False)]
        plan_view = planning[["Jour","Heure début","Heure fin","Matière","Type","Salle","Groupe"]]
        st.download_button(
            "⬇️ Exporter le planning en Excel",
            df_to_xlsx_bytes(plan_view),
            file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.xlsx",
            use_container_width=True,
        )
        st.dataframe(plan_view, use_container_width=True, hide_index=True)

    with tab_next:
        st.markdown("#### Ma prochaine séance")

        def _next(df: pd.DataFrame) -> Optional[Tuple[datetime, pd.Series]]:
            if df.empty: return None
            py = {"LUNDI":0,"MARDI":1,"MERCREDI":2,"JEUDI":3,"VENDREDI":4,"SAMEDI":5,"DIMANCHE":6}
            today = now.weekday()
            rows=[]
            for _, r in df.iterrows():
                d = py.get(str(r["Jour"]).upper(), None)
                if d is None: continue
                m = time_to_minutes(r["Heure début"]); 
                if m is None: continue
                delta = (d - today) % 7
                dt = datetime.combine((now + timedelta(days=delta)).date(), dtime.min) + timedelta(minutes=m)
                if dt < now: dt += timedelta(days=7)
                rows.append((dt, r))
            if not rows: return None
            rows.sort(key=lambda x:x[0]); return rows[0]

        nxt = _next(bloc)

        if nxt:
            dt, r = nxt
            st.markdown(
                f"""
                <div class="card">
                  <div class="next-title">{r['Matière']} <span class="badge">{r['Type']}</span></div>
                  <div class="muted">
                    {r['Jour']} • {r['Heure début']}–{r['Heure fin']}
                    <span class="badge">Salle {r['Salle']}</span>
                    <span class="badge">Groupe {r['Groupe']}</span>
                    <span class="badge">Dans {human_delta(dt, now)}</span>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("Aucune séance à venir avec ces filtres.")

    with tab_where:
        st.markdown("#### Où trouver un enseignant ?")
        only_today = st.checkbox("Aujourd’hui uniquement", value=False)

        base = edt.copy()
        base = base[base["Semestre"].astype(str).str.upper() == SEMESTRE].copy()

        if base.empty:
            st.info("Aucun cours dans les données S1.")
        else:
            base["Jour"] = base["Jour"].astype(str).str.upper().str.strip()
            base["__start"] = base["Heure début"].map(time_to_minutes)
            base["__end"]   = base["Heure fin"].map(time_to_minutes)

            today_idx = now.weekday()
            today_name = JOURS_FR[today_idx]
            now_min = now.hour*60 + now.minute

            def next_occurrence(row: pd.Series) -> Tuple[Optional[datetime], bool, bool]:
                d_idx = ORDER_JOUR.get(str(row["Jour"]).upper(), None)
                if d_idx is None or pd.isna(row["__start"]) or pd.isna(row["__end"]):
                    return None, False, False
                is_today = (d_idx == today_idx)
                if is_today and (row["__start"] <= now_min < row["__end"]):
                    dt_start = datetime.combine(now.date(), dtime.min) + timedelta(minutes=int(row["__start"]))
                    return dt_start, True, True
                delta = (d_idx - today_idx) % 7
                dt_day = (now + timedelta(days=delta)).date()
                dt_start = datetime.combine(dt_day, dtime.min) + timedelta(minutes=int(row["__start"]))
                if is_today and row["__start"] < now_min:
                    dt_start += timedelta(days=7); is_today = False
                return dt_start, is_today, False

            rows = []
            for ens, g in base.groupby("Enseignant", dropna=True):
                if not ens or str(ens).strip() == "":
                    continue
                best = None
                for _, r in g.iterrows():
                    dt, is_today, is_now = next_occurrence(r)
                    if dt is None: continue
                    cand = (dt, is_today, is_now, r)
                    if best is None or cand[0] < best[0]:
                        best = cand
                if best is None: continue

                dt, is_today, is_now, r = best
                if is_now:
                    statut = f"🟢 En cours jusqu’à {r['Heure fin']} (Salle {r['Salle']})"; order=(0,dt)
                    jour_txt = r["Jour"].title(); heure_txt=f"{r['Heure début']}–{r['Heure fin']}"
                else:
                    if is_today:
                        statut = f"🔵 Dans {human_delta(dt, now)}"; jour_txt = "Aujourd’hui"
                    else:
                        statut = f"⚪ Le {r['Jour'].title()} à {r['Heure début']} (dans {human_delta(dt, now)})"
                        jour_txt = r["Jour"].title()
                    order=(1 if is_today else 2, dt); heure_txt=f"{r['Heure début']}–{r['Heure fin']}"

                rows.append({
                    "Enseignant": ens,
                    "Statut": statut,
                    "Heure": heure_txt,
                    "Jour": jour_txt,
                    "Salle": r["Salle"],
                    "Matière": r["Matière"],
                    "Groupe": r["Groupe"],
                    "Spécialité": r.get("Spec2",""),
                    "Niveau": pretty_level_label(r.get("Spec2",""), r.get("Niv2","")),
                    "_order0": order[0], "_order1": order[1],
                })

            df_where = pd.DataFrame(rows)
            if only_today and not df_where.empty:
                df_where = df_where[df_where["Jour"].isin(["Aujourd’hui", today_name.title()])]
            if q_nom and not df_where.empty:
                df_where = df_where[df_where["Enseignant"].str.contains(q_nom, case=False, na=False)]

            if df_where.empty:
                st.info("Aucun enseignant à afficher avec les filtres actuels.")
            else:
                df_where = df_where.sort_values(by=["_order0","_order1","Enseignant"])
                st.dataframe(df_where.drop(columns=["_order0","_order1"]),
                             use_container_width=True, hide_index=True)

    # ---- Feuille de présence (enseignant) — mobile friendly + PDF
    with tab_presence:
        st.markdown("#### Feuille de présence (enseignant)")

        # ➜ mode mobile : on affiche une page de X étudiants, sans défilement horizontal
        mobile_mode = st.toggle(
            "📱 Mode mobile (affichage compact)", value=True,
            help="Affiche par pages : Nom + Présent + Remarque (idéal sur smartphone)"
        )

        q_filter = st.text_input("🔎 Recherche rapide (Nom/Prénom) :", value="").strip()

        # --- données brutes groupe
        etu_g_raw = subgroup_by_spec_level(etu, spec, niv, groupe).copy()

        # on supprime Matricule pour éviter l'encombrement
        if "Matricule" in etu_g_raw.columns:
            etu_g_raw = etu_g_raw.drop(columns=["Matricule"])

        # colonnes minimales
        if "Nom" in etu_g_raw.columns or "Prenom" in etu_g_raw.columns:
            etu_g_raw["Nom complet"] = (
                etu_g_raw.get("Nom", "").astype(str).str.strip() + " " +
                etu_g_raw.get("Prenom", "").astype(str).str.strip()
            ).str.strip().replace("^\\s+$", "", regex=True)
        else:
            first_col = etu_g_raw.columns[0] if len(etu_g_raw.columns) else "Etudiant"
            etu_g_raw["Nom complet"] = etu_g_raw[first_col].astype(str)

        # initialise l'état global (toute la liste) une seule fois
        full_key = f"presence_full_{spec}_{niv}_{groupe}"
        if full_key not in st.session_state:
            base = etu_g_raw[["Nom complet"]].dropna().drop_duplicates().reset_index(drop=True)
            base["Présent"] = False
            base["Remarque"] = ""
            st.session_state[full_key] = base

        # fusionne au cas où une nouvelle liste arrive
        full_df = st.session_state[full_key]
        incoming = etu_g_raw[["Nom complet"]].dropna().drop_duplicates().reset_index(drop=True)
        full_df = incoming.merge(full_df, on="Nom complet", how="left")
        full_df["Présent"] = full_df["Présent"].fillna(False)
        full_df["Remarque"] = full_df["Remarque"].fillna("")
        st.session_state[full_key] = full_df

        # filtre par recherche
        if q_filter:
            full_df = full_df[full_df["Nom complet"].str.contains(q_filter, case=False, na=False)]

        # pagination (évite le scroll sur smartphone)
        page_key = f"presence_page_{spec}_{niv}_{groupe}"
        if page_key not in st.session_state:
            st.session_state[page_key] = 1
        page = st.session_state[page_key]

        page_size = st.select_slider(
            "Taille de page", options=[8, 10, 12, 15, 20], value=12 if mobile_mode else 20,
            help="Nombre d'étudiants affichés simultanément"
        )
        total = len(full_df)
        max_page = max(1, (total + page_size - 1) // page_size)
        page = min(page, max_page)

        c1, c2, c3 = st.columns([1, 1, 4])
        with c1:
            if st.button("◀️ Précédent", disabled=(page <= 1), use_container_width=True):
                page = max(1, page - 1)
        with c2:
            if st.button("Suivant ▶️", disabled=(page >= max_page), use_container_width=True):
                page = min(max_page, page + 1)
        with c3:
            st.caption(f"Page {page} / {max_page} — {total} étudiants")

        st.session_state[page_key] = page
        start, end = (page - 1) * page_size, (page - 1) * page_size + page_size
        page_df = full_df.iloc[start:end].copy()

        # boutons page : (dé)cocher
        colA, colB, colC = st.columns([1, 1, 3])
        with colA:
            if st.button("✔️ Tout cocher (page)", use_container_width=True):
                page_df["Présent"] = True
        with colB:
            if st.button("✖️ Tout décocher (page)", use_container_width=True):
                page_df["Présent"] = False
        with colC:
            st.caption("En mode mobile : Nom + case Présent + Remarque, sans défilement horizontal.")

        # éditeur compact : Nom + Présent + Remarque
        col_cfg = {
            "Nom complet": st.column_config.TextColumn("Étudiant", width="large", disabled=True),
            "Présent": st.column_config.CheckboxColumn("Présent"),
            "Remarque": st.column_config.TextColumn("Remarque", width="large"),
        }

        edited_page = st.data_editor(
            page_df[["Nom complet", "Présent", "Remarque"]],
            hide_index=True,
            use_container_width=True,
            height=480 if mobile_mode else 520,
            num_rows="fixed",
            column_config=col_cfg,
            key=f"editor_presence_{spec}_{niv}_{groupe}_p{page}",
        )

        # réinjecte les modifications de la page dans la liste complète
        full_ref = st.session_state[full_key].set_index("Nom complet")
        for _, row in edited_page.iterrows():
            full_ref.loc[row["Nom complet"], "Présent"] = bool(row["Présent"])
            full_ref.loc[row["Nom complet"], "Remarque"] = str(row["Remarque"] or "")
        st.session_state[full_key] = full_ref.reset_index()

        # export PDF (avec entête UABT + date & heure)
        export_df = st.session_state[full_key].copy()
        export_df = export_df.sort_values("Nom complet").reset_index(drop=True)

        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        titre_pdf = f"Feuille de présence — {spec} {pretty_level_label(spec, niv)} • Groupe {groupe} (S1)"
        meta_pdf  = f"Généré le {ts}"

        header_pdf = {
            "univ": "UNIVERSITÉ DE TLEMCEN",
            "fac":  "FACULTÉ DE TECHNOLOGIE",
            "dept": "DÉPARTEMENT DE GÉNIE CIVIL",
            "spec": f"Spécialité : {spec} — Niveau : {pretty_level_label(spec, niv)}",
            "grp":  f"Groupe : {groupe}",
        }

        pdf_bytes = make_presence_pdf(export_df, titre_pdf, meta_pdf, header=header_pdf)
        st.download_button(
            "📄 Exporter la présence en PDF",
            data=pdf_bytes,
            file_name=f"Presence_{spec}_{niv}_G{groupe}_S1_{datetime.now():%Y%m%d_%H%M}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

# ----------------------------- FOOTER ---------------------------------

st.divider()
st.caption(
    "S1 • Spécialité → Niveau → Groupe • Groupes normalisés (G11/G12) • "
    "Harmonisation des listes étudiants • Exports EDT/Planning en Excel (.xlsx) • "
    "Feuille de présence mobile + Remarque • Export PDF officiel."
)
