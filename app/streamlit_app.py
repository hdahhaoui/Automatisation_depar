# ======================================================================
# Portail Génie Civil — EDT & Listes (S1)
# Fichier unique : app/streamlit_app.py
# ======================================================================
# Fonctionnalités :
# - Profils Étudiant / Enseignant, filtres Spécialité → Niveau → Groupe
# - Normalisation EDT & listes étudiants (S1)
# - Inférence robuste depuis nom de fichier (2ING, ING2, etc.)
# - Export Excel uniquement (.xlsx)
# - Feuille de présence mobile (sans matricule) + Tout cocher / Tout décocher
# - Panneaux diagnostic (EDT vide + Index des fichiers détectés)
# - Découverte S1 tolérante : par nom OU par contenu (“Semestre”)
#
# Arborescence attendue :
#   app/
#     streamlit_app.py
#     data/
#       raw/
#         edt/
#         students/
#
# requirements.txt :
#   streamlit
#   pandas
#   openpyxl
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

# --------------------------- CONFIG GLOBALE ----------------------------

st.set_page_config(
    page_title="EDT & Listes • Génie Civil (S1)",
    page_icon="🗓️",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
RAW_EDT = str(BASE_DIR / "data" / "raw" / "edt")
RAW_STU = str(BASE_DIR / "data" / "raw" / "students")

SEMESTRE = "S1"  # application mono-semestre

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
          h1, h2, h3 { letter-spacing:.2px }
          .actionbar { display:flex; gap:.5rem; flex-wrap:wrap; margin:.25rem 0 1rem }
          .pill { padding:.35rem .6rem; border-radius:999px; background:#1f2937; font-size:.85rem }
          .role-etudiant  { background:#0b3b2e; color:#8ef5dd }
          .role-enseignant{ background:#2a2543; color:#c3b5ff }
          .stDataFrame table { font-size: 0.92rem }
          .card { background:#0f1624; border:1px solid #1f2937; padding:1rem; border-radius:12px; margin: .25rem 0 .75rem }
          .muted{ color:#9ca3af; font-size:.9rem }
          .badge { padding:.15rem .45rem; border-radius:8px; background:#1f2937; font-size:.78rem; margin-left:.35rem }
          .sticky { position:sticky; top:0; z-index:9; background:transparent; padding-top:.25rem }
          [data-testid="stDataEditorContainer"] { border-radius: 12px; }
        </style>
        """,
        unsafe_allow_html=True,
    )

def header_role(role_label: str, subtitle: str) -> None:
    role_class = "role-etudiant" if role_label == "Étudiant" else "role-enseignant"
    st.markdown(
        f"""
        <div class="sticky">
          <div class="actionbar">
            <span class="pill {role_class}">
              {"👩‍🎓 Étudiant" if role_label == "Étudiant" else "👨‍🏫 Enseignant"}
            </span>
            <span class="pill">{subtitle}</span>
          </div>
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
    """Combine la date d et des minutes depuis minuit → datetime."""
    return datetime.combine(d.date(), dtime.min) + timedelta(minutes=minutes)

def human_delta(dt: datetime, now: datetime) -> str:
    """Délai lisible '1j 2h 30m'."""
    s = int((dt - now).total_seconds())
    d = s // 86400; s %= 86400
    h = s // 3600; s %= 3600
    m = s // 60
    out = []
    if d: out.append(f"{d}j")
    if h: out.append(f"{h}h")
    if m: out.append(f"{m}m")
    return " ".join(out) or "0m"

def next_session(now: datetime, edt_df: pd.DataFrame) -> Optional[Tuple[datetime, pd.Series]]:
    """Prochaine séance dans un EDT hebdomadaire."""
    if edt_df.empty:
        return None
    py_day = {"LUNDI":0,"MARDI":1,"MERCREDI":2,"JEUDI":3,"VENDREDI":4,"SAMEDI":5,"DIMANCHE":6}
    today_idx = now.weekday()
    rows = []
    for _, r in edt_df.iterrows():
        d_idx = py_day.get(str(r["Jour"]).upper(), None)
        if d_idx is None:
            continue
        m = time_to_minutes(r["Heure début"])
        if m is None:
            continue
        delta = (d_idx - today_idx) % 7
        dt = datetime.combine((now + timedelta(days=delta)).date(), dtime.min) + timedelta(minutes=m)
        if dt < now:
            dt += timedelta(days=7)
        rows.append((dt, r))
    if not rows:
        return None
    rows.sort(key=lambda x: x[0])
    return rows[0]

# --------------------- NORMALISATION & INFERENCE ----------------------

def ensure_cols(df: pd.DataFrame, cols: Iterable[str], numeric: Iterable[str] = ()) -> pd.DataFrame:
    """Garantit un set minimal de colonnes (valeurs vides si manquantes)."""
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
    """Détecte (Spécialité, Niveau) à partir des champs libres."""
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
    """
    Infère (Spec2, Niv2, Groupe, Semestre) depuis le nom du fichier.
    Compatible avec : 2ING, ING2, ING_2, ING-2, 2-ING, INGENIEUR2, etc.
    Groupes : G11, _G11, -G11, G 11
    Semestre : S1, _S1, S 1
    """
    name = Path(path).stem.upper()
    name_compact = re.sub(r"[\s\-]+", "_", name)  # espaces/tirets -> _
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

    # INGENIEUR 1/2/3 (toutes formes)
    m = re.search(r"(?:INGENIEUR|ING)[_\- ]?([123])", name_compact)
    if not m:
        m = re.search(r"([123])[_\- ]?(?:INGENIEUR|ING)", name_compact)
    if m:
        return "INGENIEUR", m.group(1), g, (sem or "S1")

    # Par défaut
    return None, None, g, (sem or "S1")


# ------------------ HARMONISATION LISTES ÉTUDIANTS --------------------

def harmonize_student_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Renomme les colonnes variées vers le schéma standard + split Nom/Prénom si besoin."""
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
    """
    Charge TOUS les fichiers, puis garde S1 si :
      - le NOM du fichier contient 'S1' (insensible à la casse), OU
      - la COLONNE 'Semestre' contient 'S1'.
    Retourne: (edt, etu, idx) où idx = index de fichiers détectés (diagnostic).
    """
    idx_rows = []  # lignes diagnostic

    # ----------------------------- EDT ---------------------------------
    edt_frames = []
    for f in sorted(glob.glob(f"{RAW_EDT}/*.*")):
        fname = Path(f).name
        name_up = fname.upper()

        try:
            df = read_any(f)
        except Exception as e:
            idx_rows.append({
                "Type": "EDT", "Fichier": fname, "Lu": False,
                "Erreur": str(e), "S1_par_nom": "S1" in name_up, "S1_par_col": None,
                "Spec2": None, "Niv2": None, "Groupe": None
            })
            continue

        df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
        df["Semestre"] = df["Semestre"].apply(normalize_semestre)
        df["Groupe"] = df["Groupe"].apply(normalize_groupe)

        has_s1_col = df["Semestre"].astype(str).str.upper().eq("S1").any()
        s1_by_name = "S1" in name_up
        if not (s1_by_name or has_s1_col):
            idx_rows.append({
                "Type": "EDT", "Fichier": fname, "Lu": True,
                "Erreur": "", "S1_par_nom": s1_by_name, "S1_par_col": has_s1_col,
                "Spec2": None, "Niv2": None, "Groupe": None
            })
            continue

        # --- Classification initiale depuis colonnes libres
        specs, nivs = zip(*df.apply(
            lambda r: classify_spec_level(r.get("Spécialité",""), r.get("Niveau","")),
            axis=1
        ))
        df["Spec2"], df["Niv2"] = specs, nivs

        # --- INFÉRENCE + FORÇAGE GLOBAL depuis NOM DE FICHIER (PATCH)
        s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
        df["Spec2"]   = df["Spec2"].fillna("").astype(str).str.upper()
        df["Niv2"]    = df["Niv2"].fillna("").astype(str).str.upper()
        df["Groupe"]  = df["Groupe"].fillna("").astype(str)
        df["Semestre"]= df["Semestre"].fillna("").astype(str).str.upper()
        if s2_f:  df["Spec2"]    = s2_f
        if n2_f:  df["Niv2"]     = n2_f
        if g_f:   df["Groupe"]   = normalize_groupe(g_f)
        if sem_f: df["Semestre"] = sem_f

        # --- Index diag
        idx_rows.append({
            "Type": "EDT",
            "Fichier": fname,
            "Lu": True,
            "Erreur": "",
            "S1_par_nom": s1_by_name,
            "S1_par_col": has_s1_col,
            "Spec2": df["Spec2"].dropna().astype(str).str.upper().replace("", None).mode().tolist()[:1] or [None],
            "Niv2": df["Niv2"].dropna().astype(str).str.upper().replace("", None).mode().tolist()[:1] or [None],
            "Groupe": df["Groupe"].dropna().map(normalize_groupe).mode().tolist()[:1] or [None],
        })

        # Ne garde que S1
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
            idx_rows.append({
                "Type": "ETU", "Fichier": fname, "Lu": False,
                "Erreur": str(e), "S1_par_nom": "S1" in name_up, "S1_par_col": None,
                "Spec2": None, "Niv2": None, "Groupe": None
            })
            continue

        df = harmonize_student_columns(df)
        df = ensure_cols(df, STU_COLS)
        df["Semestre"] = df["Semestre"].apply(normalize_semestre)
        df["Groupe"] = df["Groupe"].apply(normalize_groupe)

        has_s1_col = df["Semestre"].astype(str).str.upper().eq("S1").any()
        s1_by_name = "S1" in name_up
        if not (s1_by_name or has_s1_col):
            idx_rows.append({
                "Type": "ETU", "Fichier": fname, "Lu": True,
                "Erreur": "", "S1_par_nom": s1_by_name, "S1_par_col": has_s1_col,
                "Spec2": None, "Niv2": None, "Groupe": None
            })
            continue

        specs, nivs = zip(*df.apply(
            lambda r: classify_spec_level(r.get("Spécialité",""), r.get("Niveau","")),
            axis=1
        ))
        df["Spec2"], df["Niv2"] = specs, nivs

        # --- INFÉRENCE + FORÇAGE GLOBAL depuis NOM DE FICHIER (PATCH)
        s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
        df["Spec2"]   = df["Spec2"].fillna("").astype(str).str.upper()
        df["Niv2"]    = df["Niv2"].fillna("").astype(str).str.upper()
        df["Groupe"]  = df["Groupe"].fillna("").astype(str)
        df["Semestre"]= df["Semestre"].fillna("").astype(str).str.upper()
        if s2_f:  df["Spec2"]    = s2_f
        if n2_f:  df["Niv2"]     = n2_f
        if g_f:   df["Groupe"]   = normalize_groupe(g_f)
        if sem_f: df["Semestre"] = sem_f

        idx_rows.append({
            "Type": "ETU",
            "Fichier": fname,
            "Lu": True,
            "Erreur": "",
            "S1_par_nom": s1_by_name,
            "S1_par_col": has_s1_col,
            "Spec2": df["Spec2"].dropna().astype(str).str.upper().replace("", None).mode().tolist()[:1] or [None],
            "Niv2": df["Niv2"].dropna().astype(str).str.upper().replace("", None).mode().tolist()[:1] or [None],
            "Groupe": df["Groupe"].dropna().map(normalize_groupe).mode().tolist()[:1] or [None],
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
# ============================ INTERFACE ===============================

st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu, idx = load_raw_s1()

# ---- Index de détection (diagnostic complet)
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

# 🔍 Diagnostic si EDT vide pour les filtres actuels
if bloc.empty:
    with st.expander("🔍 Diagnostic (EDT vide pour ce filtre)", expanded=False):
        df_spec = edt[edt["Spec2"] == spec]
        st.write("Niveaux disponibles pour", spec, ":", sorted(df_spec["Niv2"].dropna().unique().tolist()))
        st.write("Groupes vus pour", spec, niv, ":", sorted(
            df_spec[df_spec["Niv2"] == niv]["Groupe"].dropna().map(normalize_groupe).unique().tolist()
        ))
        st.write("Exemples de fichiers détectés dans EDT :", [Path(p).name for p in glob.glob(f"{RAW_EDT}/*")][:12])

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
        nxt = next_session(now, bloc)
        if nxt:
            dt, r = nxt
            st.markdown(
                f"""
                <div class="card">
                  <div style="font-size:1.1rem;font-weight:600">{r['Matière']} <span class="badge">{r['Type']}</span></div>
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
        nxt = next_session(now, bloc)
        if nxt:
            dt, r = nxt
            st.markdown(
                f"""
                <div class="card">
                  <div style="font-size:1.1rem;font-weight:600">{r['Matière']} <span class="badge">{r['Type']}</span></div>
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

    # ---- Feuille de présence (enseignant) — version mobile friendly
    with tab_presence:
        st.markdown("#### Feuille de présence (enseignant)")

        # Affichage compact pour smartphone
        mobile_mode = st.toggle("📱 Mode mobile (affichage compact)", value=True,
                                help="Affiche seulement Nom et Présent (idéal sur smartphone)")

        # Recherche rapide
        q_filter = st.text_input("🔎 Recherche rapide (Nom/Prénom) :", value="").strip()

        # Charge la liste harmonisée pour le groupe
        etu_g_raw = subgroup_by_spec_level(etu, spec, niv, groupe).copy()

        # Supprimer Matricule s’il existe (écran + export)
        if "Matricule" in etu_g_raw.columns:
            etu_g_raw = etu_g_raw.drop(columns=["Matricule"])

        # Colonnes de base : privilégier N°, Nom, Prenom, Remarque
        base_cols_pref = ["N°", "Nom", "Prenom", "Remarque"]
        base_cols = [c for c in base_cols_pref if c in etu_g_raw.columns]
        if not base_cols:
            base_cols = [c for c in etu_g_raw.columns if c not in {"Spec2","Niv2","Semestre"}][:3]

        etu_g = etu_g_raw[base_cols].reset_index(drop=True)

        if etu_g.empty:
            st.warning("Pas de liste trouvée pour ce groupe (vérifie 'Groupe' = G11/G12 et 'Semestre' = S1).")
        else:
            # Colonne Nom complet
            if "Nom" in etu_g.columns or "Prenom" in etu_g.columns:
                etu_g["Nom complet"] = (etu_g.get("Nom","").astype(str).str.strip() + " " +
                                        etu_g.get("Prenom","").astype(str).str.strip()).str.strip()
            else:
                first_col = etu_g.columns[0]
                etu_g["Nom complet"] = etu_g[first_col].astype(str)

            # Ajoute "Présent" si absente
            if "Présent" not in etu_g.columns:
                etu_g["Présent"] = False

            # Filtre texte
            if q_filter:
                mask = etu_g["Nom complet"].str.contains(q_filter, case=False, na=False)
                if "Remarque" in etu_g.columns:
                    mask = mask | etu_g["Remarque"].astype(str).str.contains(q_filter, case=False, na=False)
                etu_g = etu_g[mask].reset_index(drop=True)

            # Jeu de colonnes final selon le mode
            if mobile_mode:
                show_cols = ["Nom complet", "Présent"]
            else:
                order = []
                if "N°" in etu_g.columns: order.append("N°")
                order += ["Nom complet"]
                if "Remarque" in etu_g.columns: order.append("Remarque")
                show_cols = [c for c in order if c in etu_g.columns] + ["Présent"]

            # État en session (conserve cases cochées)
            key_df = f"presence_{spec}_{niv}_{groupe}"
            if key_df not in st.session_state:
                st.session_state[key_df] = etu_g.copy()[show_cols]
            else:
                # Ré-aligner si colonnes changent (toggle mobile)
                missing = [c for c in show_cols if c not in st.session_state[key_df].columns]
                if missing:
                    st.session_state[key_df] = etu_g.copy()[show_cols]
                else:
                    # Conserver Présent par "Nom complet"
                    left = etu_g.copy()[show_cols]
                    if "Nom complet" in show_cols:
                        prev = st.session_state[key_df].set_index("Nom complet")
                        new  = left.set_index("Nom complet")
                        if "Présent" in prev.columns and "Présent" in new.columns:
                            new["Présent"] = new.index.map(prev["Présent"]).fillna(False)
                        st.session_state[key_df] = new.reset_index()
                    else:
                        st.session_state[key_df] = left

            # Actions rapides
            colA, colB, colC = st.columns([1,1,2])
            with colA:
                if st.button("✔️ Tout cocher", use_container_width=True):
                    st.session_state[key_df]["Présent"] = True
            with colB:
                if st.button("✖️ Tout décocher", use_container_width=True):
                    st.session_state[key_df]["Présent"] = False
            with colC:
                st.caption("En mode mobile : Nom + case Présent seulement (pas de défilement horizontal).")

            # Config colonnes
            col_cfg = {
                "Nom complet": st.column_config.TextColumn("Étudiant", width="large", disabled=True),
            }
            if "N°" in show_cols:
                col_cfg["N°"] = st.column_config.NumberColumn("N°", width="small", disabled=True)
            if "Remarque" in show_cols:
                col_cfg["Remarque"] = st.column_config.TextColumn("Remarque", width="medium")
            if "Présent" in show_cols:
                col_cfg["Présent"] = st.column_config.CheckboxColumn("Présent", help="Cocher la présence")

            edited = st.data_editor(
                st.session_state[key_df],
                column_config=col_cfg,
                hide_index=True,
                use_container_width=True,
                height=520,
                num_rows="fixed",
                key=f"editor_{key_df}_{'m' if mobile_mode else 'd'}",
            )
            st.session_state[key_df] = edited

            # Export Excel (sans Matricule)
            st.download_button(
                "⬇️ Exporter la présence en Excel",
                df_to_xlsx_bytes(edited),
                file_name=f"Presence_{spec}_{niv}_G{groupe}_S1.xlsx",
                use_container_width=True,
            )

            # CSS lisibilité mobile
            st.markdown("""
                <style>
                [data-testid="stDataEditorRow"] { min-height: 40px; }
                [data-testid="column-Nom complet"] div { white-space: normal !important; }
                </style>
            """, unsafe_allow_html=True)

# ----------------------------- FOOTER ---------------------------------

st.divider()
st.caption(
    "S1 • Spécialité → Niveau → Groupe • Groupes normalisés (G11/G12) • "
    "Harmonisation des listes étudiants • Exports uniquement en Excel (.xlsx) • "
    "Feuille de présence ergonomique sur smartphone (sans Matricule)."
)
