# app/streamlit_app.py — S1, filtres hiérarchiques Spécialité → Niveau, libellés propres
import glob
from io import BytesIO
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------- Config ----------------
st.set_page_config(page_title="EDT & Listes • Génie Civil (S1)", page_icon="🗓️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent            # → /app
RAW_EDT = str(BASE_DIR / "data" / "raw" / "edt")
RAW_STU = str(BASE_DIR / "data" / "raw" / "students")
SEMESTRE = "S1"                                       # un seul semestre

ORDER_JOUR = {"DIMANCHE":0,"LUNDI":1,"MARDI":2,"MERCREDI":3,"JEUDI":4,"VENDREDI":5,"SAMEDI":6}

# ---------------- Utils ----------------
def read_any(path):
    return pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)

def df_to_xlsx_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
    return buf.getvalue()

def time_to_minutes(h):
    h = str(h).strip().lower().replace(" ", "")
    if "h" not in h: return None
    hh, mm = h.split("h")
    return int(hh or 0)*60 + int(mm or 0)

def human_delta(dt, now):
    s = int((dt - now).total_seconds())
    d = s // 86400; s %= 86400
    h = s // 3600;  s %= 3600
    m = s // 60
    out = []
    if d: out.append(f"{d}j")
    if h: out.append(f"{h}h")
    if m: out.append(f"{m}m")
    return " ".join(out) or "0m"

def next_session(now, edt_df):
    if edt_df.empty: return None
    py_day = {"LUNDI":0,"MARDI":1,"MERCREDI":2,"JEUDI":3,"VENDREDI":4,"SAMEDI":5,"DIMANCHE":6}
    today_idx = now.weekday()
    rows = []
    for _, r in edt_df.iterrows():
        d_idx = py_day.get(str(r["Jour"]).upper(), None)
        if d_idx is None: continue
        delta = (d_idx - today_idx) % 7
        day_date = (now + timedelta(days=delta)).date()
        m = time_to_minutes(r["Heure début"])
        if m is None: continue
        dt = datetime.combine(day_date, datetime.min.time()) + timedelta(minutes=m)
        if dt < now:
            dt = dt + timedelta(days=7)
        rows.append((dt, r))
    if not rows: return None
    rows.sort(key=lambda x:x[0])
    return rows[0][0], rows[0][1]

# ---------------- Normalisation colonnes ----------------
EDT_COLS = ["Niveau","Spécialité","Groupe","Semestre","Jour","Heure début","Heure fin",
            "Durée (h)","Matière","Type","Enseignant","Salle","Fréquence"]
STU_COLS = ["Annee","Semestre","Spécialité","Niveau","Groupe","Matricule","Nom","Prenom",
            "Email","Téléphone","Remarque","N°"]

def ensure_cols(df, cols, numeric=None):
    numeric = set(numeric or [])
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0 if c in numeric else ""
    return df[cols]

# ---------------- Classification Spécialité/Niveau ----------------
def classify_spec_level(spec_text: str, level_text: str):
    """
    Retourne (Spec2, Niv2) normalisés:
      Spec2 ∈ {"RIB","VOA","STRUCTURE","LICENCE","INGENIEUR"}
      Niv2  ∈ {"M1","M2","2","3","1","2","3"} selon la spécialité
    Règles:
      - Masters: RIB/VOA/STRUCTURE -> M1 ou M2
      - Licence: L2/L3 -> LICENCE + 2/3
      - Ingénieur: ING -> INGENIEUR + 1/2/3
    """
    S = (spec_text or "").upper()
    L = (level_text or "").upper()

    # Masters d'abord
    if "RIB" in S:
        niv = "M1" if ("M1" in S or "M1" in L) else ("M2" if ("M2" in S or "M2" in L) else "")
        return "RIB", niv
    if "VOA" in S:
        niv = "M1" if ("M1" in S or "M1" in L) else ("M2" if ("M2" in S or "M2" in L) else "")
        return "VOA", niv
    if "STRUCT" in S:
        niv = "M1" if ("M1" in S or "M1" in L) else ("M2" if ("M2" in S or "M2" in L) else "")
        return "STRUCTURE", niv

    # Licence
    if "L2" in S or "L2" in L or "LICENCE 2" in S:
        return "LICENCE", "2"
    if "L3" in S or "L3" in L or "LICENCE 3" in S:
        return "LICENCE", "3"

    # Ingénieur
    if "ING" in S or "INGÉ" in S or "INGENIEUR" in S or "INGÉNIEUR" in S or "ING" in L:
        if "1" in S or "1" in L: return "INGENIEUR", "1"
        if "2" in S or "2" in L: return "INGENIEUR", "2"
        if "3" in S or "3" in L: return "INGENIEUR", "3"
        return "INGENIEUR", ""

    # fallback
    return "", ""

def level_options_for(spec: str):
    if spec in ("RIB","VOA","STRUCTURE"): return ["M1","M2"]
    if spec == "LICENCE": return ["2","3"]
    if spec == "INGENIEUR": return ["1","2","3"]
    return []

def pretty_level_label(spec: str, niv: str):
    if spec == "LICENCE": return f"LICENCE {niv}"
    if spec == "INGENIEUR": return f"INGENIEUR {niv}"
    return niv  # Masters : M1/M2

# ---------------- Chargement (depuis data/raw/*) ----------------
@st.cache_data
def load_raw_s1():
    # EDT
    edt_files = glob.glob(f"{RAW_EDT}/*_S1.*")
    edt_list = []
    for f in edt_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
            df["Semestre"] = df["Semestre"].replace("", SEMESTRE)
            df["Fréquence"] = df["Fréquence"].replace("", "Hebdo")
            sp2, lv2 = [], []
            for _, r in df.iterrows():
                s2, n2 = classify_spec_level(r.get("Spécialité",""), r.get("Niveau",""))
                sp2.append(s2); lv2.append(n2)
            df["Spec2"] = sp2
            df["Niv2"]  = lv2
            edt_list.append(df)
        except Exception as e:
            st.warning(f"EDT ignoré: {f} ({e})")
    edt = pd.concat(edt_list, ignore_index=True) if edt_list else pd.DataFrame(columns=EDT_COLS+["Spec2","Niv2"])
    edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
    edt = edt.sort_values(["Spec2","Niv2","Groupe","__o","Heure début"]).drop(columns="__o")

    # Étudiants
    stu_files = glob.glob(f"{RAW_STU}/*_S1.*")
    stu_list = []
    for f in stu_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, STU_COLS)
            sp2, lv2 = [], []
            for _, r in df.iterrows():
                s2, n2 = classify_spec_level(r.get("Spécialité",""), r.get("Niveau",""))
                sp2.append(s2); lv2.append(n2)
            df["Spec2"] = sp2
            df["Niv2"]  = lv2
            stu_list.append(df)
        except Exception as e:
            st.warning(f"Liste ignorée: {f} ({e})")
    etu = pd.concat(stu_list, ignore_index=True) if stu_list else pd.DataFrame(columns=STU_COLS+["Spec2","Niv2"])
    return edt, etu

def subgroup_by_spec_level(df, spec=None, niv=None, groupe=None):
    keep = df[df["Semestre"].astype(str).str.upper()==SEMESTRE]
    if spec:   keep = keep[keep["Spec2"]==spec]
    if niv:    keep = keep[keep["Niv2"]==niv]
    if groupe: keep = keep[keep["Groupe"].astype(str).str.upper()==str(groupe).upper()]
    return keep

# ---------------- UI ----------------
st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu = load_raw_s1()
if edt.empty:
    st.error("Aucun EDT S1 trouvé dans `app/data/raw/edt/` (fichiers `_S1.xlsx`/`.csv`).")
    st.stop()
if etu.empty:
    st.warning("Aucune liste d’étudiants S1 trouvée dans `app/data/raw/students/` (l’appli fonctionne mais sans feuilles de présence).")

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)

    st.markdown("---")
    st.caption("Filtres hiérarchiques")

    # 1) Spécialité (ordre imposé)
    spec_order = ["RIB","VOA","STRUCTURE","LICENCE","INGENIEUR"]
    available_specs = [s for s in spec_order if s in edt["Spec2"].dropna().unique().tolist()]
    spec = st.selectbox("Sp
