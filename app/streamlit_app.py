# app/streamlit_app.py — S1 uniquement, charge depuis data/raw/*
import glob
from io import BytesIO
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="EDT & Listes • Génie Civil (S1)", page_icon="🗓️", layout="wide")

RAW_EDT = "data/raw/edt"
RAW_STU = "data/raw/students"
ORDER_JOUR = {"DIMANCHE":0,"LUNDI":1,"MARDI":2,"MERCREDI":3,"JEUDI":4,"VENDREDI":5,"SAMEDI":6}
SEMESTRE = "S1"     # <— un seul semestre

# ---------- utils lecture ----------
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
        if dt < now:  # déjà passé aujourd'hui
            dt = dt + timedelta(days=7)
        rows.append((dt, r))
    if not rows: return None
    rows.sort(key=lambda x:x[0])
    return rows[0][0], rows[0][1]

# ---------- normalisation ----------
EDT_COLS = ["Niveau","Spécialité","Groupe","Semestre","Jour","Heure début","Heure fin","Durée (h)","Matière","Type","Enseignant","Salle","Fréquence"]
STU_COLS = ["Annee","Semestre","Spécialité","Niveau","Groupe","Matricule","Nom","Prenom","Email","Téléphone","Remarque","N°"]

def ensure_cols(df, cols, numeric=None):
    numeric = set(numeric or [])
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0 if c in numeric else ""
    return df[cols]

@st.cache_data
def load_raw_s1():
    # EDT
    edt_files = glob.glob(f"{RAW_EDT}/*_S1.*")
    edt_list = []
    for f in edt_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
            # Forcer S1 si vide
            df["Semestre"] = df["Semestre"].replace("", SEMESTRE)
            # Valeurs par défaut
            df["Fréquence"] = df["Fréquence"].replace("", "Hebdo")
            edt_list.append(df)
        except Exception as e:
            st.warning(f"EDT ignoré: {f} ({e})")
    edt = pd.concat(edt_list, ignore_index=True) if edt_list else pd.DataFrame(columns=EDT_COLS)
    edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
    edt = edt.sort_values(["Niveau","Spécialité","Groupe","__o","Heure début"]).drop(columns="__o")

    # Étudiants
    stu_files = glob.glob(f"{RAW_STU}/*_S1.*")
    stu_list = []
    for f in stu_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, STU_COLS)
            stu_list.append(df)
        except Exception as e:
            st.warning(f"Liste ignorée: {f} ({e})")
    etu = pd.concat(stu_list, ignore_index=True) if stu_list else pd.DataFrame(columns=STU_COLS)

    return edt, etu

def subgroup(df, niveau=None, spec=None, groupe=None):
    keep = df[df["Semestre"].astype(str).str.upper()==SEMESTRE]
    if niveau:   keep = keep[keep["Niveau"].astype(str).str.upper()==niveau.upper()]
    if spec:     keep = keep[keep["Spécialité"].astype(str).str.contains(spec, case=False, na=False)]
    if groupe:   keep = keep[keep["Groupe"].astype(str).str.upper()==groupe.upper()]
    return keep

# ---------- UI ----------
st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu = load_raw_s1()
if edt.empty:
    st.error("Aucun EDT S1 trouvé dans `data/raw/edt/`. Vérifie que tes fichiers se terminent par `_S1.xlsx` ou `_S1.csv` et sont committés dans le dépôt.")
    st.stop()
if etu.empty:
    st.warning("Aucune liste d’étudiants S1 trouvée dans `data/raw/students/` (l’appli fonctionne mais sans feuilles de présence).")

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)
    st.caption("Filtres")
    niveau  = st.selectbox("Niveau", sorted(edt["Niveau"].dropna().unique()), index=None, placeholder="Tous")
    spec    = st.selectbox("Spécialité", sorted(edt["Spécialité"].dropna().unique()), index=None, placeholder="Toutes")
    groupe  = st.selectbox("Groupe", sorted(edt["Groupe"].dropna().unique()), index=None, placeholder="Tous")
    st.caption("Recherche")
    q_nom   = st.text_input("Nom/Prénom (étudiant ou enseignant)")
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

bloc = subgroup(edt, niveau, spec, groupe)
now = datetime.now()

# ---- Étudiant ----
if role == "Étudiant":
    st.header("👩‍🎓 Espace Étudiant")

    col1, col2 = st.columns([2,1])
    with col1:
        st.subheader("Mon EDT")
        view = bloc[["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence"]]
        st.dataframe(view, use_container_width=True)
        st.download_button("⬇️ Export (CSV)", view.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"EDT_{(niveau or 'X')}_{(groupe or 'X')}_S1.csv")
        st.download_button("⬇️ Export (XLSX)", df_to_xlsx_bytes(view),
                           file_name=f"EDT_{(niveau or 'X')}_{(groupe or 'X')}_S1.xlsx")
    with col2:
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

    st.subheader("Liste des étudiants (groupe sélectionné)")
    etu_g = subgroup(etu, niveau, spec, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Pas de liste trouvée pour ce groupe.")
    else:
        etu_g["Présent"] = False
        edited = st.data_editor(etu_g, use_container_width=True, height=420, num_rows="fixed")
        st.download_button("⬇️ Feuille de présence (CSV)",
                           edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"presence_{(niveau or 'X')}_{(groupe or 'X')}_S1.csv")

    st.subheader("Je cherche mon nom")
    if q_nom:
        hits = etu[etu.apply(lambda r: q_nom.lower() in f"{r['Nom']} {r['Prenom']}".lower(), axis=1)]
        st.dataframe(hits[["Nom","Prenom","Niveau","Spécialité","Groupe","Semestre","Remarque"]], use_container_width=True)

# ---- Enseignant ----
else:
    st.header("👨‍🏫 Espace Enseignant")

    planning = bloc.copy()
    if q_nom:
        planning = planning[planning["Enseignant"].str.contains(q_nom, case=False, na=False)]
    planning_view = planning[["Jour","Heure début","Heure fin","Matière","Type","Groupe","Salle","Fréquence","Spécialité"]]
    st.subheader("Planning filtré")
    st.dataframe(planning_view, use_container_width=True)
    st.download_button("⬇️ Export planning (CSV)", planning_view.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"planning_{(q_nom or 'enseignant')}_S1.csv")
    st.download_button("⬇️ Export planning (XLSX)", df_to_xlsx_bytes(planning_view),
                       file_name=f"planning_{(q_nom or 'enseignant')}_S1.xlsx")

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

    st.subheader("Trouver une salle pour un enseignant")
    if q_nom:
        salles = edt[edt["Enseignant"].str.contains(q_nom, case=False, na=False)][["Jour","Heure début","Heure fin","Salle","Groupe","Spécialité"]]
        st.dataframe(salles.sort_values(["Jour","Heure début"]), use_container_width=True)

    st.subheader("Liste des étudiants du groupe sélectionné")
    etu_g = subgroup(etu, niveau, spec, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Pas de liste pour ce groupe.")
    else:
        etu_g["Présent"] = False
        edited = st.data_editor(etu_g, use_container_width=True, height=420, num_rows="fixed")
        st.download_button("⬇️ Feuille de présence (CSV)",
                           edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"presence_{(niveau or 'X')}_{(groupe or 'X')}_S1.csv")

st.divider()
st.caption("S1 uniquement • Lecture directe de data/raw/edt & data/raw/students • Sessions isolées (Streamlit).")
