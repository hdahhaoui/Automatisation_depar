# app/streamlit_app.py
import time
from io import BytesIO
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

# ---------------- Config ----------------
st.set_page_config(page_title="EDT & Listes • Génie Civil", page_icon="🗓️", layout="wide")

DATA_DIR = "data/processed"          # dossiers générés par build_master.py
ORDER_JOUR = {"DIMANCHE":0,"LUNDI":1,"MARDI":2,"MERCREDI":3,"JEUDI":4,"VENDREDI":5,"SAMEDI":6}
SEMESTRES = ["S1","S2"]              # prêt pour S2

# --------------- Utils ---------------
def file_for(prefix, sem):
    return f"{DATA_DIR}/{prefix}_MASTER_{sem}.xlsx"

@st.cache_data
def load_data(semestre):
    edt_file = file_for("EDT", semestre)
    etu_file = file_for("ETUDIANTS", semestre)
    edt = pd.read_excel(edt_file)
    etu = pd.read_excel(etu_file)
    # nettoyage minimal
    for col in ["Niveau","Spécialité","Groupe","Semestre","Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence","Durée (h)"]:
        if col not in edt.columns:
            edt[col] = "" if col!="Durée (h)" else 0.0
    for col in ["Annee","Semestre","Spécialité","Niveau","Groupe","Nom","Prenom","Matricule","N°","Remarque"]:
        if col not in etu.columns:
            etu[col] = ""
    # tri lisible
    edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
    edt = edt.sort_values(["Niveau","Spécialité","Groupe","__o","Heure début"]).drop(columns="__o")
    return edt, etu

def time_to_minutes(h):
    h = str(h).strip().lower().replace(" ", "")
    if "h" not in h: return None
    hh, mm = h.split("h")
    return int(hh or 0)*60 + int(mm or 0)

def next_session(now, edt_df, jour_col="Jour"):
    if edt_df.empty: return None
    py_day = {"LUNDI":0,"MARDI":1,"MERCREDI":2,"JEUDI":3,"VENDREDI":4,"SAMEDI":5,"DIMANCHE":6}
    today_idx = now.weekday()
    rows = []
    for _, r in edt_df.iterrows():
        d_idx = py_day.get(str(r[jour_col]).upper(), None)
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

def subgroup(df, semestre="S1", niveau=None, spec=None, groupe=None):
    keep = df.copy()
    if semestre: keep = keep[keep["Semestre"].astype(str).str.upper()==semestre.upper()]
    if niveau:   keep = keep[keep["Niveau"].astype(str).str.upper()==niveau.upper()]
    if spec:     keep = keep[keep["Spécialité"].astype(str).str.contains(spec, case=False, na=False)]
    if groupe:   keep = keep[keep["Groupe"].astype(str).str.upper()==groupe.upper()]
    return keep

def df_to_xlsx_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
    return buf.getvalue()

# --------------- UI : sidebar ---------------
with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)
    semestre = st.selectbox("Semestre", SEMESTRES, index=0)
    st.markdown("---")
    st.caption("Filtres rapides")
    # chargement initial pour remplir les options
    _edt, _etu = load_data(semestre)
    niveau  = st.selectbox("Niveau", sorted(_edt["Niveau"].dropna().unique()), index=None, placeholder="Tous")
    spec    = st.selectbox("Spécialité", sorted(_edt["Spécialité"].dropna().unique()), index=None, placeholder="Toutes")
    groupe  = st.selectbox("Groupe", sorted(_edt["Groupe"].dropna().unique()), index=None, placeholder="Tous")
    st.markdown("---")
    st.caption("Recherche")
    q_nom   = st.text_input("Nom/Prénom (étudiant ou enseignant)")
    st.markdown("---")
    print_mode = st.checkbox("🖨️ Mode impression (masquer la barre latérale)")

# --------------- Impression : CSS ---------------
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

# --------------- Chargement des données ---------------
edt, etu = load_data(semestre)
now = datetime.now()
bloc = subgroup(edt, semestre, niveau, spec, groupe)

st.title("🗓️ Portail Génie Civil — EDT & Listes")

# ------------------- Étudiant -------------------
if role == "Étudiant":
    st.header("👩‍🎓 Espace Étudiant")

    col1, col2 = st.columns([2,1])
    with col1:
        st.subheader("Mon EDT")
        view = bloc[["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence"]]
        st.dataframe(view, use_container_width=True)
        st.download_button("⬇️ Export (CSV)", data=view.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"EDT_{(niveau or 'X')}_{(groupe or 'X')}_{semestre}.csv",
                           mime="text/csv")
        st.download_button("⬇️ Export (XLSX)", data=df_to_xlsx_bytes(view),
                           file_name=f"EDT_{(niveau or 'X')}_{(groupe or 'X')}_{semestre}.xlsx")

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
            st.info("Aucune séance trouvée avec les filtres actuels.")

    st.subheader("Liste des étudiants (groupe sélectionné)")
    etu_g = subgroup(etu, semestre, niveau, spec, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Aucune liste d’étudiants correspondante.")
    else:
        etu_g["Présent"] = False
        edited = st.data_editor(etu_g, use_container_width=True, height=420, num_rows="fixed")
        st.download_button("⬇️ Feuille de présence (CSV)",
                           data=edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"presence_{(niveau or 'X')}_{(groupe or 'X')}_{semestre}.csv",
                           mime="text/csv")
        st.caption("ℹ️ Coche Présent/Absent, puis télécharge le CSV ou fais une capture (Ctrl/Cmd+P).")

    st.subheader("Je cherche mon nom")
    if q_nom:
        hits = etu[etu.apply(lambda r: q_nom.lower() in f"{r['Nom']} {r['Prenom']}".lower(), axis=1)]
        st.dataframe(hits[["Nom","Prenom","Niveau","Spécialité","Groupe","Semestre","Remarque"]], use_container_width=True)

# ------------------- Enseignant -------------------
else:
    st.header("👨‍🏫 Espace Enseignant")

    # Planning filtré & export
    planning = bloc.copy()
    if q_nom:
        planning = planning[planning["Enseignant"].str.contains(q_nom, case=False, na=False)]
    planning_view = planning[["Jour","Heure début","Heure fin","Matière","Type","Groupe","Salle","Fréquence","Spécialité"]]
    st.subheader("Planning filtré")
    st.dataframe(planning_view, use_container_width=True)
    st.download_button("⬇️ Export planning (CSV)", data=planning_view.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"planning_{(q_nom or 'enseignant')}_{semestre}.csv", mime="text/csv")
    st.download_button("⬇️ Export planning (XLSX)", data=df_to_xlsx_bytes(planning_view),
                       file_name=f"planning_{(q_nom or 'enseignant')}_{semestre}.xlsx")

    # Prochaine séance + compte à rebours
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
        st.info("Aucune séance trouvée avec les filtres actuels.")

    # Où trouver un enseignant (salles)
    st.subheader("Trouver une salle pour un enseignant")
    if q_nom:
        salles = edt[edt["Enseignant"].str.contains(q_nom, case=False, na=False)][["Jour","Heure début","Heure fin","Salle","Groupe","Spécialité"]]
        st.dataframe(salles.sort_values(["Jour","Heure début"]), use_container_width=True)

    # Liste des étudiants du groupe
    st.subheader("Liste des étudiants du groupe sélectionné")
    etu_g = subgroup(etu, semestre, niveau, spec, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Aucune liste d’étudiants correspondante.")
    else:
        etu_g["Présent"] = False
        edited = st.data_editor(etu_g, use_container_width=True, height=420, num_rows="fixed")
        st.download_button("⬇️ Feuille de présence (CSV)",
                           data=edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"presence_{(niveau or 'X')}_{(groupe or 'X')}_{semestre}.csv",
                           mime="text/csv")

    # --------- CHARGE HORAIRE ---------
    st.subheader("Charge horaire (heures / semaine)")
    # somme des Durée (h), groupée
    charges = edt.groupby(["Enseignant","Matière"], dropna=False)["Durée (h)"].sum().reset_index()
    if q_nom:
        charges = charges[charges["Enseignant"].str.contains(q_nom, case=False, na=False)]
    # total par enseignant
    total_ens = charges.groupby("Enseignant", dropna=False)["Durée (h)"].sum().reset_index().rename(columns={"Durée (h)":"Total (h)"})
    st.write("**Total par enseignant :**")
    st.dataframe(total_ens.sort_values("Total (h)", ascending=False), use_container_width=True, height=260)
    st.download_button("⬇️ Export charge totale (CSV)", data=total_ens.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"charge_totale_{semestre}.csv", mime="text/csv")
    st.write("**Détail par matière :**")
    st.dataframe(charges.sort_values(["Enseignant","Durée (h)"], ascending=[True,False]), use_container_width=True, height=300)
    st.download_button("⬇️ Export charge détaillée (CSV)", data=charges.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"charge_detail_{semestre}.csv", mime="text/csv")

st.divider()
st.caption("Données : EDT_MASTER_* & ETUDIANTS_MASTER_* • Sessions isolées (Streamlit) • Mode impression via Ctrl/Cmd+P.")
