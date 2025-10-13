# ======================================================================
# Portail Génie Civil — EDT & Listes (S1)
# Fichier : app/streamlit_app.py
# ======================================================================
# - Profils Étudiant / Enseignant
# - Filtres Spécialité → Niveau → Groupe
# - Normalisation EDT & listes étudiants (S1)
# - Inférence robuste depuis nom de fichier (ING2/2ING, etc.)
# - Exports Excel (.xlsx) pour EDT/Planning
# - Feuille de présence mobile (sans matricule) + Remarque + Tout cocher/décocher
# - Export PDF présence avec en-tête institutionnel UABT
# - Prochaine séance (partagé) + bandeau coloré + sélecteur de jour
# - Planning : filtre par Salle
# - Annuaire enseignants (hebdomadaire + salles)
# ======================================================================

from __future__ import annotations

import re
import glob
from io import BytesIO
from datetime import datetime, timedelta, time as dtime, time
from pathlib import Path
from typing import Tuple, Optional, Iterable, Dict, Any

import pandas as pd
import streamlit as st

# --- PDF (présence)
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors

# --- Timezone pour décomptes & “aujourd’hui”
import pytz

# --- Utilitaires Prochaine séance (Algérie / parsing) ---
import datetime as _dt
import pytz as _pytz

TZ_DZ = _pytz.timezone("Africa/Algiers")
# mapping Python weekday() -> libellé FR utilisé dans tes tableaux
WEEKDAY_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]

def now_dz():
    """datetime 'aware' en Afrique/Alger."""
    return _dt.datetime.now(TZ_DZ)

def parse_hhmm(s: str) -> _dt.time:
    """
    '08h30' -> datetime.time(8,30)
    Accepte aussi '08:30' par sécurité.
    """
    if not isinstance(s, str):
        s = str(s or "")
    s = s.strip().lower().replace(" ", "")
    s = s.replace("h", ":")
    if ":" not in s:
        # ex: '8' -> '8:00'
        s = f"{s}:00"
    hh, mm = s.split(":", 1)
    return _dt.time(int(hh), int(mm or 0))

def enrich_times(df):
    """Ajoute colonnes _tstart/_tend (time) et tri par Heure début."""
    if df.empty:
        return df.copy()
    out = df.copy()
    out["_tstart"] = out["Heure début"].apply(parse_hhmm)
    out["_tend"]   = out["Heure fin"].apply(parse_hhmm)
    out = out.sort_values(by=["_tstart", "_tend"]).reset_index(drop=True)
    return out

def pick_today_label():
    """Libellé de jour FR (ex: 'MERCREDI') basé sur l’heure algérienne."""
    return WEEKDAY_FR[now_dz().weekday()]

def pick_current_and_next(sessions_for_day):
    """
    sessions_for_day: DataFrame déjà enrichi (colonnes _tstart/_tend).
    Retourne (current_row | None, next_row | None, state_str)
      - state_str ∈ {'ongoing','upcoming','empty','completed'}
    """
    if sessions_for_day.empty:
        return None, None, "empty"

    tnow = now_dz().time()
    # séance en cours ?
    ongoing = sessions_for_day[
        (sessions_for_day["_tstart"] <= tnow) &
        (tnow < sessions_for_day["_tend"])
    ]
    if not ongoing.empty:
        cur = ongoing.iloc[0]
        # la prochaine après celle en cours
        nexts = sessions_for_day[sessions_for_day["_tstart"] > cur["_tend"]]
        nxt = nexts.iloc[0] if not nexts.empty else None
        return cur, nxt, "ongoing"

    # pas de séance en cours: cherche la prochaine
    nxts = sessions_for_day[sessions_for_day["_tstart"] > tnow]
    if not nxts.empty:
        nxt = nxts.iloc[0]
        return None, nxt, "upcoming"

    # toutes les séances du jour sont passées
    return None, None, "completed"

def fmt_hhmm(t: _dt.time) -> str:
    return f"{t.hour:02d}h{t.minute:02d}"

def td_to_hm(delta: _dt.timedelta) -> str:
    # Retour "XhYY" ou "X min" si < 1h
    secs = int(delta.total_seconds())
    if secs < 0:
        secs = 0
    h, r = divmod(secs, 3600)
    m, _ = divmod(r, 60)
    if h == 0:
        return f"{m} min"
    return f"{h}h{m:02d}"


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
SHOW_DIAGNOSTIC = False           # <- True pour afficher l’expander diagnostic

# Jour / ordre
WEEKDAY_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]
JOURS_FR = WEEKDAY_FR
ORDER_JOUR = {d: i for i, d in enumerate(WEEKDAY_FR)}

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

          /* (Optionnel) Classes pour les cartes de séance si tu veux éviter le inline style */
          .session-card{border:1px solid var(--secondary-background-color); border-radius:10px; overflow:hidden; margin-bottom:8px;}
          .session-head{padding:.55rem .8rem; font-weight:700; color:#fff;}
          .session-body{padding:.55rem .8rem;}
          .session-foot{padding:0 .8rem .7rem .8rem; opacity:.9;}
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
    if "h" not in s and ":" not in s:
        return None
    try:
        s = s.replace("h", ":")
        hh, mm = s.split(":")
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

        # --- INFÉRENCE + FORÇAGE GLOBAL
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

        # --- INFÉRENCE + FORÇAGE GLOBAL
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
    spec_line = header.get("spec", "").strip()
    grp_line  = header.get("grp", "").strip()

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
# ====================== UTILS AVANCÉS (Salles/Profs/Horaires) ======================

def tz_now():
    try:
        tz = pytz.timezone("Africa/Algiers")
        return datetime.now(tz)
    except Exception:
        return datetime.now()

def french_weekday_name(dt=None):
    if dt is None: dt = tz_now()
    return WEEKDAY_FR[dt.weekday()]  # Lundi=0

def _to_dt_today(hour_str, base_date=None):
    if base_date is None:
        base_date = tz_now().date()
    s = str(hour_str).replace("h", ":")
    hh, mm = s.split(":")
    hh, mm = int(hh), int(mm)
    tz = getattr(tz_now(), "tzinfo", None)
    return datetime.combine(base_date, time(hh, mm, 0, tzinfo=tz))

def _col(df, name_candidates):
    cols = {c.lower(): c for c in df.columns}
    for cand in name_candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    raise KeyError(f"Colonne manquante parmi {name_candidates} dans {list(df.columns)}")

def get_cols(df):
    return dict(
        jour=_col(df, ["Jour"]),
        hdeb=_col(df, ["Heure début","Début","Heure_debut","Heure Debut"]),
        hfin=_col(df, ["Heure fin","Fin","Heure_fin","Heure Fin"]),
        mat=_col(df, ["Matière","Matiere","Module"]),
        typ=_col(df, ["Type"]),
        ens=_col(df, ["Enseignant","Prof","Intervenant"]),
        sal=_col(df, ["Salle"]),
        freq=_col(df, ["Fréquence","Frequence"]),
        grp=_col(df, ["Groupe"])
    )

def current_and_next_for_day(df):
    if df is None or df.empty:
        return (None, None, None)
    C = get_cols(df)
    today = french_weekday_name()
    dday = df[df[C["jour"]].astype(str).str.upper().eq(today)].copy()
    if dday.empty:
        return (None, None, None)
    dday["_dt_start"] = dday[C["hdeb"]].astype(str).apply(_to_dt_today)
    dday["_dt_end"]   = dday[C["hfin"]].astype(str).apply(_to_dt_today)
    dday.sort_values("_dt_start", inplace=True)
    now = tz_now()
    cur = dday[(dday["_dt_start"] <= now) & (now < dday["_dt_end"])].head(1)
    nxt = dday[dday["_dt_start"] > now].head(1)
    nxt2 = dday[dday["_dt_start"] > now].iloc[1:2]
    return (
        None if cur.empty else cur.iloc[0],
        None if nxt.empty else nxt.iloc[0],
        None if nxt2.empty else nxt2.iloc[0],
    )

def unique_rooms(df):
    if df is None or df.empty: return []
    C = get_cols(df)
    return sorted([s for s in df[C["sal"]].dropna().astype(str).unique() if s.strip()])

def unique_teachers(df_all):
    if df_all is None or df_all.empty: return []
    C = get_cols(df_all)
    return sorted([e for e in df_all[C["ens"]].dropna().astype(str).unique() if e.strip()])

def week_schedule_for_teacher(df_all, teacher):
    if not teacher or df_all is None or df_all.empty:
        return None
    C = get_cols(df_all)
    sub = df_all[df_all[C["ens"]].astype(str).str.fullmatch(teacher, case=False, na=False)].copy()
    if sub.empty: return sub
    day_index = {d:i for i,d in enumerate(WEEKDAY_FR)}
    sub["_didx"] = sub[C["jour"]].astype(str).str.upper().map(day_index).fillna(7).astype(int)
    sub["_tstart"] = sub[C["hdeb"]].astype(str).apply(lambda s: int(s.replace("h",":").replace(":","")[:4]))
    sub.sort_values(["_didx","_tstart"], inplace=True)
    return sub

# ========== Onglet Prochaine séance ==========
with tabs_area[1]:  # adapte si ton nom de variable diffère (ex: tab2)
    st.subheader("Prochaine séance (Étudiant)")

    # On filtre d’abord EDT selon spécialité/niveau/groupe déjà choisis dans ta sidebar
    edt_filtered = filtered_edt_df  # réutilise ta DataFrame déjà filtrée (spéc→niv→groupe)
    if edt_filtered is None or edt_filtered.empty:
        st.info("Aucun EDT n'est disponible pour ces filtres.")
        st.stop()

    # Ajoute colonnes temps et normalise l’EDT
    edt_enriched = enrich_times(edt_filtered)

    # Jour par défaut = jour actuel (Algérie)
    default_day = pick_today_label()
    all_days = WEEKDAY_FR  # ou edt_enriched['Jour'].str.upper().unique() trié si tu veux limiter aux jours présents
    day_choice = st.selectbox("Jour", options=all_days, index=all_days.index(default_day))

    # Séances du jour choisi
    day_df = edt_enriched[edt_enriched["Jour"].str.upper() == day_choice].copy()

    cur, nxt, state = pick_current_and_next(day_df)

    # helpers d’affichage
    def block_session(row, headline="Séance", color="#4F7BFE"):
        mat = str(row["Matière"])
        typ = str(row.get("Type", "") or "").strip()
        teach = str(row.get("Enseignant", "") or "")
        salle = str(row.get("Salle", "") or "")
        t1, t2 = row["_tstart"], row["_tend"]

        st.markdown(
            f"""
            <div style="background:{color};color:white;border-radius:10px;padding:10px 14px;margin-top:8px;">
              <strong>{mat} ({typ})</strong><br/>
              👨‍🏫 {teach} &nbsp; • &nbsp; 🏫 Salle {salle} &nbsp; • &nbsp; 📅 {day_choice}<br/>
              🕒 {fmt_hhmm(t1)} – {fmt_hhmm(t2)}
            </div>
            """,
            unsafe_allow_html=True
        )

    # Affichage logique
    st.caption("Prochaine séance")

    now_local = now_dz()
    if state == "empty":
        st.warning(f"Aucune séance planifiée pour {day_choice}.")
    elif state == "ongoing":
        # séance en cours
        block_session(cur, headline="Séance en cours", color="#16a34a")  # vert
        # temps restant
        end_dt = now_local.replace(hour=cur["_tend"].hour, minute=cur["_tend"].minute, second=0, microsecond=0)
        left = end_dt - now_local
        st.markdown(f"⏱ **En cours** — reste **{td_to_hm(left)}** (de {fmt_hhmm(cur['_tstart'])} à {fmt_hhmm(cur['_tend'])}).")

        st.markdown("**Après :**")
        if nxt is not None:
            block_session(nxt, headline="Prochaine", color="#4F7BFE")
            start_dt = now_local.replace(hour=nxt["_tstart"].hour, minute=nxt["_tstart"].minute, second=0, microsecond=0)
            st.caption(f"🗓 Dans **{td_to_hm(start_dt - now_local)}** (début {fmt_hhmm(nxt['_tstart'])}).")
        else:
            st.info("Aucune autre séance après celle en cours.")
    elif state == "upcoming":
        # prochaine séance à venir
        block_session(nxt, headline="Prochaine", color="#4F7BFE")
        start_dt = now_local.replace(hour=nxt["_tstart"].hour, minute=nxt["_tstart"].minute, second=0, microsecond=0)
        st.caption(f"🗓 Dans **{td_to_hm(start_dt - now_local)}** (de {fmt_hhmm(nxt['_tstart'])} à {fmt_hhmm(nxt['_tend'])}).")
        # Bonus: montre aussi la suivante après celle-là
        rest = day_df[day_df["_tstart"] > nxt["_tstart"]]
        if not rest.empty:
            st.markdown("**Après :**")
            nxt2 = rest.iloc[0]
            block_session(nxt2, headline="Ensuite", color="#7c3aed")  # violet
    else:  # completed
        st.info(f"Toutes les séances de **{day_choice}** sont terminées.")

# ============================ INTERFACE ===============================

st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu, idx = load_raw_s1()
df_all_edt = edt.copy()  # EDT global S1 pour l'annuaire

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
        # Composant partagé (jour sélectionnable + bandeau coloré + enseignant)
        render_next_sessions_shared(bloc, title="Prochaine séance (Étudiant)")
# =========================== VUE ENSEIGNANT ===========================

else:
    header_role("Enseignant", f"{title_clean} • Groupe {groupe}")

    tab_plan, tab_next, tab_annuaire, tab_presence = st.tabs(
        ["🗂️ Planning", "⏭️ Prochaine séance", "📇 Annuaire enseignants", "📝 Feuille de présence"]
    )

    # ---------- Planning avec filtre par Salle ----------
    with tab_plan:
        st.markdown("#### Planning filtré")
        planning = bloc.copy()
        if not planning.empty:
            # filtre nom enseignant libre (optionnel)
            if q_nom:
                C = get_cols(planning)
                planning = planning[planning[C["ens"]].str.contains(q_nom, case=False, na=False)]

            C = get_cols(planning)
            salles_opts = ["Toutes"] + unique_rooms(planning)
            sel_salle = st.selectbox("Filtrer par salle", salles_opts, index=0, key="filtre_salle_planning")
            bloc_aff = planning.copy()
            if sel_salle != "Toutes":
                bloc_aff = bloc_aff[bloc_aff[C["sal"]].astype(str) == sel_salle]

            # Ordonner visuellement
            try:
                day_index = {d:i for i,d in enumerate(WEEKDAY_FR)}
                bloc_aff["_didx"] = bloc_aff[C["jour"]].astype(str).str.upper().map(day_index).fillna(7).astype(int)
                bloc_aff["_tstart"] = bloc_aff[C["hdeb"]].astype(str).apply(lambda s: int(s.replace("h",":").replace(":","")[:4]))
                bloc_aff = bloc_aff.sort_values(["_didx","_tstart"]).drop(columns=["_didx","_tstart"])
            except Exception:
                pass

            plan_view = bloc_aff[[C["jour"],C["hdeb"],C["hfin"],C["mat"],C["typ"],C["sal"],C["grp"]]].rename(columns={
                C["jour"]: "Jour", C["hdeb"]: "Début", C["hfin"]: "Fin", C["mat"]: "Matière", C["typ"]: "Type",
                C["sal"]: "Salle", C["grp"]: "Groupe"
            })

            st.download_button(
                "⬇️ Exporter le planning en Excel",
                df_to_xlsx_bytes(plan_view),
                file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.xlsx",
                use_container_width=True,
            )
            st.dataframe(plan_view, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun cours pour ce filtre.")

    # ---------- Prochaine séance (partagé) ----------
    with tab_next:
        render_next_sessions_shared(bloc, title="Prochaine séance (Enseignant)")

    # ---------- Annuaire enseignants ----------
    with tab_annuaire:
        st.markdown("#### Annuaire des enseignants (hebdomadaire)")
        if df_all_edt.empty:
            st.info("Aucun EDT global S1 disponible.")
        else:
            q_t = st.text_input("Recherche par nom", key="search_teacher").strip()
            teachers = unique_teachers(df_all_edt)
            if q_t:
                teachers = [t for t in teachers if q_t.lower() in t.lower()]

            if not teachers:
                st.info("Aucun enseignant trouvé pour ce filtre.")
            else:
                teach = st.selectbox("Sélectionnez un enseignant", ["— choisir —"] + teachers, index=0, key="sb_teacher_dir")
                if teach and teach != "— choisir —":
                    sub = week_schedule_for_teacher(df_all_edt, teach)
                    if sub is None or sub.empty:
                        st.warning("Aucun créneau pour cet enseignant.")
                    else:
                        C = get_cols(sub)
                        salles = sorted(s for s in sub[C["sal"]].dropna().astype(str).unique() if s.strip())
                        st.caption("Salles hebdomadaires : " + (" • ".join([f"`{s}`" for s in salles]) if salles else "—"))
                        show = sub[[C["jour"], C["hdeb"], C["hfin"], C["mat"], C["typ"], C["sal"], C["grp"]]].rename(columns={
                            C["jour"]: "Jour", C["hdeb"]: "Début", C["hfin"]: "Fin", C["mat"]: "Matière", C["typ"]: "Type",
                            C["sal"]: "Salle", C["grp"]: "Groupe"
                        })
                        st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.write("### Liste des enseignants")
                    cols = st.columns(3)
                    for i, t in enumerate(teachers):
                        cols[i % 3].write(f"- {t}")

    # ---------- Feuille de présence (enseignant) ----------
    with tab_presence:
        st.markdown("#### Feuille de présence (enseignant)")

        mobile_mode = st.toggle(
            "📱 Mode mobile (affichage compact)", value=True,
            help="Affiche par pages : Nom + Présent + Remarque (idéal sur smartphone)"
        )

        q_filter = st.text_input("🔎 Recherche rapide (Nom/Prénom) :", value="").strip()

        etu_g_raw = subgroup_by_spec_level(etu, spec, niv, groupe).copy()

        if "Matricule" in etu_g_raw.columns:
            etu_g_raw = etu_g_raw.drop(columns=["Matricule"])

        if "Nom" in etu_g_raw.columns or "Prenom" in etu_g_raw.columns:
            etu_g_raw["Nom complet"] = (
                etu_g_raw.get("Nom", "").astype(str).str.strip() + " " +
                etu_g_raw.get("Prenom", "").astype(str).str.strip()
            ).str.strip().replace("^\\s+$", "", regex=True)
        else:
            first_col = etu_g_raw.columns[0] if len(etu_g_raw.columns) else "Etudiant"
            etu_g_raw["Nom complet"] = etu_g_raw[first_col].astype(str)

        full_key = f"presence_full_{spec}_{niv}_{groupe}"
        if full_key not in st.session_state:
            base = etu_g_raw[["Nom complet"]].dropna().drop_duplicates().reset_index(drop=True)
            base["Présent"] = False
            base["Remarque"] = ""
            st.session_state[full_key] = base

        full_df = st.session_state[full_key]
        incoming = etu_g_raw[["Nom complet"]].dropna().drop_duplicates().reset_index(drop=True)
        full_df = incoming.merge(full_df, on="Nom complet", how="left")
        full_df["Présent"] = full_df["Présent"].fillna(False)
        full_df["Remarque"] = full_df["Remarque"].fillna("")
        st.session_state[full_key] = full_df

        if q_filter:
            full_df = full_df[full_df["Nom complet"].str.contains(q_filter, case=False, na=False)]

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

        colA, colB, colC = st.columns([1, 1, 3])
        with colA:
            if st.button("✔️ Tout cocher (page)", use_container_width=True):
                page_df["Présent"] = True
        with colB:
            if st.button("✖️ Tout décocher (page)", use_container_width=True):
                page_df["Présent"] = False
        with colC:
            st.caption("En mode mobile : Nom + case Présent + Remarque, sans défilement horizontal.")

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

        full_ref = st.session_state[full_key].set_index("Nom complet")
        for _, row in edited_page.iterrows():
            full_ref.loc[row["Nom complet"], "Présent"] = bool(row["Présent"])
            full_ref.loc[row["Nom complet"], "Remarque"] = str(row["Remarque"] or "")
        st.session_state[full_key] = full_ref.reset_index()

        export_df = st.session_state[full_key].copy().sort_values("Nom complet").reset_index(drop=True)

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
    "Prochaine séance partagée (jour sélectionnable) • Annuaire enseignants • "
    "Feuille de présence mobile + Remarque • Export PDF officiel (UABT)."
)
