# app/streamlit_app.py — S1 • Spécialité→Niveau • Groupes robustes • Détection par nom de fichier
import re
import glob
from io import BytesIO
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------- Config ----------------
st.set_page_config(page_title="EDT & Listes • Génie Civil (S1)", page_icon="🗓️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
RAW_EDT = str(BASE_DIR / "data" / "raw" / "edt")
RAW_STU = str(BASE_DIR / "data" / "raw" / "students")
SEMESTRE = "S1"

ORDER_JOUR = {"DIMANCHE":0,"LUNDI":1,"MARDI":2,"MERCREDI":3,"JEUDI":4,"VENDREDI":5,"SAMEDI":6}

# ---------------- Utils génériques ----------------
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

# ---------------- Normalisation & classification ----------------
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

def normalize_semestre(val, fallback="S1"):
    v = str(val).strip().upper()
    return v or fallback

def normalize_groupe(val):
    s = str(val).upper().replace(" ", "")
    if s and not s.startswith("G"):
        if s.isdigit():
            s = "G" + s
    return s

def classify_spec_level(spec_text: str, level_text: str):
    S = (spec_text or "").upper()
    L = (level_text or "").upper()

    # Masters
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
    if "L2" in S or "L2" in L or "LICENCE 2" in S: return "LICENCE", "2"
    if "L3" in S or "L3" in L or "LICENCE 3" in S: return "LICENCE", "3"

    # Ingénieur
    if "ING" in S or "INGÉ" in S or "INGENIEUR" in S or "INGÉNIEUR" in S or "ING" in L:
        if "1" in S or "1" in L: return "INGENIEUR", "1"
        if "2" in S or "2" in L: return "INGENIEUR", "2"
        if "3" in S or "3" in L: return "INGENIEUR", "3"
        return "INGENIEUR", ""
    return "", ""

def infer_from_filename(path: str):
    """
    Essaye d'inférer (Spec2, Niv2, Groupe, Semestre) depuis le nom du fichier.
    Ex: ETUDIANTS_M1_RIB_G11_S1.xlsx → ("RIB","M1","G11","S1")
        ETUDIANTS_1ING_G12_S1.xlsx   → ("INGENIEUR","1","G12","S1")
        ETUDIANTS_L3_G11_S1.xlsx     → ("LICENCE","3","G11","S1")
    """
    name = Path(path).stem.upper().replace("-", "_")
    # Groupe
    g = None
    m = re.search(r"_G\s*?(\d+)", name)
    if m: g = f"G{m.group(1)}"

    # Semestre
    sem = None
    m = re.search(r"_S\s*?(\d+)", name)
    if m: sem = f"S{m.group(1)}"

    # Masters
    if "RIB" in name:   return "RIB", ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "VOA" in name:   return "VOA", ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "STRUCT" in name:return "STRUCTURE", ("M2" if "M2" in name else "M1"), g, (sem or "S1")

    # Licence
    if "L2" in name:    return "LICENCE", "2", g, (sem or "S1")
    if "L3" in name:    return "LICENCE", "3", g, (sem or "S1")

    # Ingénieur
    if "1ING" in name:  return "INGENIEUR", "1", g, (sem or "S1")
    if "2ING" in name:  return "INGENIEUR", "2", g, (sem or "S1")
    if "3ING" in name:  return "INGENIEUR", "3", g, (sem or "S1")

    return None, None, g, (sem or "S1")

def level_options_for(spec: str):
    if spec in ("RIB","VOA","STRUCTURE"): return ["M1","M2"]
    if spec == "LICENCE": return ["2","3"]
    if spec == "INGENIEUR": return ["1","2","3"]
    return []

def pretty_level_label(spec: str, niv: str):
    if spec == "LICENCE": return f"LICENCE {niv}"
    if spec == "INGENIEUR": return f"INGENIEUR {niv}"
    return niv  # Masters M1/M2

# ---------------- Chargement ----------------
@st.cache_data
def load_raw_s1():
    # ---- EDT
    edt_files = glob.glob(f"{RAW_EDT}/*_S1.*")
    edt_list = []
    for f in edt_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
            # normalisation de base
            df["Semestre"] = df["Semestre"].apply(normalize_semestre)
            df["Groupe"]   = df["Groupe"].apply(normalize_groupe)
            # classification
            sp2, lv2 = [], []
            for _, r in df.iterrows():
                s2, n2 = classify_spec_level(r.get("Spécialité",""), r.get("Niveau",""))
                sp2.append(s2); lv2.append(n2)
            df["Spec2"], df["Niv2"] = sp2, lv2
            edt_list.append(df)
        except Exception as e:
            st.warning(f"EDT ignoré: {f} ({e})")
    edt = pd.concat(edt_list, ignore_index=True) if edt_list else pd.DataFrame(columns=EDT_COLS+["Spec2","Niv2"])
    edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
    edt = edt.sort_values(["Spec2","Niv2","Groupe","__o","Heure début"]).drop(columns="__o")

    # ---- Étudiants
    stu_files = glob.glob(f"{RAW_STU}/*_S1.*")
    stu_list = []
    for f in stu_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, STU_COLS)
            # normalisation de base
            df["Semestre"] = df["Semestre"].apply(normalize_semestre)
            df["Groupe"]   = df["Groupe"].apply(normalize_groupe)
            # classification via colonnes
            sp2, lv2 = [], []
            for _, r in df.iterrows():
                s2, n2 = classify_spec_level(r.get("Spécialité",""), r.get("Niveau",""))
                sp2.append(s2); lv2.append(n2)
            df["Spec2"], df["Niv2"] = sp2, lv2

            # fallback: si Spec2/Niv2/Groupe manquent, essayer depuis le nom du fichier
            if (df["Spec2"] == "").any() or (df["Niv2"] == "").any() or (df["Groupe"] == "").any():
                s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
                if s2_f:
                    df.loc[df["Spec2"]=="", "Spec2"] = s2_f
                if n2_f:
                    df.loc[df["Niv2"]=="", "Niv2"] = n2_f
                if g_f:
                    df.loc[df["Groupe"]=="", "Groupe"] = normalize_groupe(g_f)
                if sem_f:
                    df.loc[df["Semestre"]=="", "Semestre"] = sem_f

            stu_list.append(df)
        except Exception as e:
            st.warning(f"Liste ignorée: {f} ({e})")
    etu = pd.concat(stu_list, ignore_index=True) if stu_list else pd.DataFrame(columns=STU_COLS+["Spec2","Niv2"])
    return edt, etu

def subgroup_by_spec_level(df, spec=None, niv=None, groupe=None):
    keep = df[df["Semestre"].astype(str).str.upper()==SEMESTRE]
    if spec:   keep = keep[keep["Spec2"]==spec]
    if niv:    keep = keep[keep["Niv2"]==niv]
    if groupe:
        gnorm = normalize_groupe(groupe)
        keep = keep[keep["Groupe"].apply(normalize_groupe)==gnorm]
    return keep

# ---------------- UI ----------------
st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu = load_raw_s1()
if edt.empty:
    st.error("Aucun EDT S1 trouvé dans app/data/raw/edt/")
    st.stop()

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)

    st.markdown("---")
    st.caption("Filtres hiérarchiques")

    spec_order = ["RIB","VOA","STRUCTURE","LICENCE","INGENIEUR"]
    available_specs = [s for s in spec_order if s in edt["Spec2"].dropna().unique().tolist() or s in etu["Spec2"].dropna().unique().tolist()]
    spec = st.selectbox("Spécialité", available_specs, index=0 if available_specs else None)

    raw_levels = level_options_for(spec)
    level_labels = [pretty_level_label(spec, n) for n in raw_levels]
    label_to_raw = dict(zip(level_labels, raw_levels))
    niv_label = st.selectbox("Niveau", level_labels, index=0 if level_labels else None)
    niv = label_to_raw.get(niv_label)

    # Groupes provenant d'abord de l'EDT, sinon fallback sur listes d'étudiants
    g_from_edt = subgroup_by_spec_level(edt, spec, niv)["Groupe"].dropna().map(normalize_groupe)
    g_from_etu = subgroup_by_spec_level(etu, spec, niv)["Groupe"].dropna().map(normalize_groupe)
    grp_pool = sorted(pd.concat([g_from_edt, g_from_etu]).unique().tolist())
    groupe = st.selectbox("Groupe", grp_pool, index=0 if grp_pool else None)

    st.markdown("---")
    q_nom = st.text_input("Nom/Prénom (étudiant ou enseignant)")
    print_mode = st.checkbox("🖨️ Mode impression")

if print_mode:
    st.markdown("""
        <style>
            section[data-testid="stSidebar"] { display: none !important; }
            .block-container { padding-top: 1rem; padding-bottom: 0; }
            header { visibility: hidden; height: 0; }
            @media print {
                .stButton, .stDownloadButton, [data-testid="stFileUploader"] { display: none !important; }
                .stDataFrame { border: none !important; }
            }
        </style>
    """, unsafe_allow_html=True)

bloc = subgroup_by_spec_level(edt, spec, niv, groupe)
now = datetime.now()
title_clean = f"{spec} {pretty_level_label(spec, niv)}".strip()

# ---------------- Espace Étudiant ----------------
if role == "Étudiant":
    st.header(f"👩‍🎓 Espace Étudiant — {title_clean}")

    st.subheader(f"EDT — {title_clean} • Groupe {groupe}")
    view = bloc[["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence"]]
    st.dataframe(view, use_container_width=True)
    st.download_button("⬇️ Export EDT (CSV)", view.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"EDT_{spec}_{niv}_G{groupe}_S1.csv")
    st.download_button("⬇️ Export EDT (XLSX)", df_to_xlsx_bytes(view),
                       file_name=f"EDT_{spec}_{niv}_G{groupe}_S1.xlsx")

    st.subheader("Prochaine séance")
    nxt = next_session(now, bloc)
    if nxt:
        dt, r = nxt
        st.metric(
            label=f"{r['Jour']} • {r['Heure début']}–{r['Heure fin']}",
            value=f"{r['Matière']} ({r['Type']})",
            delta=f"Dans {human_delta(dt, now)} • Salle {r['Salle']} • {r['Enseignant']}"
        )
    else:
        st.info("Aucune séance avec ces filtres.")

    st.subheader("Liste des étudiants")
    etu_g = subgroup_by_spec_level(etu, spec, niv, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Pas de liste trouvée pour ce couple Spécialité/Niveau/Groupe. Vérifie la colonne 'Groupe' (G11/G12) et 'Semestre' (S1) du fichier étudiants.")
    else:
        etu_g["Présent"] = False
        st.data_editor(etu_g, use_container_width=True, height=420)

# ---------------- Espace Enseignant ----------------
else:
    st.header(f"👨‍🏫 Espace Enseignant — {title_clean}")

    st.subheader(f"Planning — {title_clean} • Groupe {groupe}")
    planning = bloc.copy()
    if q_nom:
        planning = planning[planning["Enseignant"].str.contains(q_nom, case=False, na=False)]
    st.dataframe(planning[["Jour","Heure début","Heure fin","Matière","Type","Salle","Groupe"]], use_container_width=True)
    st.download_button("⬇️ Export Planning (CSV)", planning.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.csv")

    st.subheader("Ma prochaine séance")
    nxt = next_session(now, planning)
    if nxt:
        dt, r = nxt
        st.metric(
            label=f"{r['Jour']} • {r['Heure début']}–{r['Heure fin']}",
            value=f"{r['Matière']} ({r['Type']})",
            delta=f"Dans {human_delta(dt, now)} • Salle {r['Salle']} • Groupe {r['Groupe']}"
        )
    else:
        st.info("Aucune séance avec ces filtres.")

    st.subheader("Où trouver un enseignant ?")
    if q_nom:
        salles = subgroup_by_spec_level(edt, spec, niv, None)
        salles = salles[salles["Enseignant"].str.contains(q_nom, case=False, na=False)][["Jour","Heure début","Heure fin","Salle","Groupe","Matière"]]
        st.dataframe(salles.sort_values(["Jour","Heure début"]), use_container_width=True)

    st.subheader("Liste des étudiants (groupe sélectionné)")
    etu_g = subgroup_by_spec_level(etu, spec, niv, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Pas de liste trouvée pour ce groupe.")
    else:
        etu_g["Présent"] = False
        st.data_editor(etu_g, use_container_width=True, height=420)

st.divider()
st.caption("S1 • Spécialité → Niveau → Groupe • Groupes normalisés (G11/G12) • Détection de Spec/Niv/Groupe par nom de fichier si besoin.")
