from pathlib import Path
from typing import List, Optional, Tuple
import unicodedata

import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# ---------------- Config ----------------
st.set_page_config(page_title="EDT & Listes • Génie Civil", page_icon="🗓️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DATA_ROOT = BASE_DIR.parent / "data"
PREFERRED_DATA_DIR = DATA_ROOT / "processed"  # dossiers générés par build_master.py
EDT_FILENAME = "EDT_MASTER_S1.xlsx"
ETU_FILENAME = "ETUDIANTS_MASTER_S1.xlsx"

ORDER_JOUR = {"DIMANCHE":0,"LUNDI":1,"MARDI":2,"MERCREDI":3,"JEUDI":4,"VENDREDI":5,"SAMEDI":6}

# --------------- Utils ---------------
def _normalize_filename(name: str) -> str:
    """Normalise un nom pour des comparaisons tolérantes (espaces/accents/ponctuation)."""
    normalized = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    return (
        normalized.upper()
        .replace("_", "")
        .replace("-", "")
        .replace(" ", "")
    )


def _locate_data_file(filename: str) -> Tuple[Optional[Path], List[Path]]:
    """Retourne le premier fichier correspondant trouvé et la liste des dossiers inspectés."""
    searched_dirs = []
    target_norm = _normalize_filename(filename)

    def _check_dir(directory: Path) -> Optional[Path]:
        if not directory or not directory.exists() or not directory.is_dir():
            return None
        searched_dirs.append(directory)

        candidate = directory / filename
        if candidate.exists():
            return candidate

        # tolère les variantes de casse ou les noms avec espaces/underscores différents
        for file in directory.glob("*.xls*"):
            if _normalize_filename(file.name) == target_norm:
                return file
        return None

    # 1) dossiers privilégiés
    for directory in [PREFERRED_DATA_DIR, DATA_ROOT, DATA_ROOT / "raw"]:
        found = _check_dir(directory)
        if found:
            return found, searched_dirs

    # 2) sous-dossiers directs de data/raw (ex: edt/, students/)
    raw_dir = DATA_ROOT / "raw"
    if raw_dir.exists():
        for child in sorted(raw_dir.iterdir(), key=lambda p: p.name.lower()):
            if child.is_dir():
                found = _check_dir(child)
                if found:
                    return found, searched_dirs

    # 3) recherche exhaustive dans data/
    if DATA_ROOT.exists():
        searched_dirs.append(DATA_ROOT)
        for match in sorted(DATA_ROOT.rglob("*.xls*"), key=lambda p: str(p)):
            if _normalize_filename(match.name) == target_norm:
                return match, searched_dirs

    return None, searched_dirs


@st.cache_data
def load_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    edt_cols = [
        "Niveau",
        "Spécialité",
        "Groupe",
        "Semestre",
        "Jour",
        "Heure début",
        "Heure fin",
        "Matière",
        "Type",
        "Enseignant",
        "Salle",
        "Fréquence",
    ]
    etu_cols = [
        "Annee",
        "Semestre",
        "Spécialité",
        "Niveau",
        "Groupe",
        "Nom",
        "Prenom",
        "Matricule",
        "N°",
        "Remarque",
    ]
    edt_path, edt_dirs = _locate_data_file(EDT_FILENAME)
    etu_path, etu_dirs = _locate_data_file(ETU_FILENAME)

    if not edt_path or not etu_path:
        searched = set(edt_dirs + etu_dirs)
        if not searched and DATA_ROOT.exists():
            searched = {DATA_ROOT}
        pretty_paths = []
        for p in sorted(searched, key=lambda x: str(x)):
            try:
                pretty_paths.append(str(p.relative_to(BASE_DIR.parent)))
            except ValueError:
                pretty_paths.append(str(p))
        searched_msg = ", ".join(pretty_paths) or "data/"
        st.error(
            "Les fichiers sources sont introuvables. Vérifie que le dossier `data/` contient bien les fichiers Excel attendus ("
            f"{EDT_FILENAME} & {ETU_FILENAME}).\nDossiers vérifiés : {searched_msg}."
        )
        return pd.DataFrame(columns=edt_cols), pd.DataFrame(columns=etu_cols)

    edt = pd.read_excel(edt_path)
    etu = pd.read_excel(etu_path)
    # nettoyage minimal
    for col in edt_cols:
        if col not in edt.columns:
            edt[col] = ""
    for col in etu_cols:
        if col not in etu.columns:
            etu[col] = ""
    # tri lisible
    edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
    edt = edt.sort_values(["Niveau","Spécialité","Groupe","__o","Heure début"]).drop(columns="__o")
    return edt, etu

def time_to_minutes(h):
    # '08h30' -> minutes
    h = str(h).strip().lower().replace(" ", "")
    if "h" not in h: return None
    hh, mm = h.split("h")
    return int(hh)*60 + int(mm or 0)

def next_session(now, edt_df, jour_col="Jour"):
    """Retourne la prochaine séance >= maintenant (jour+heure), sinon la plus proche du prochain jour."""
    if edt_df.empty: return None
    # construire datetime cible pour chaque ligne (semaine en cours)
    today_idx = now.weekday()  # 0=lundi ... 6=dimanche
    # On mappe nos jours à Python : DIMANCHE=6, LUNDI=0 ...
    py_day = {"LUNDI":0,"MARDI":1,"MERCREDI":2,"JEUDI":3,"VENDREDI":4,"SAMEDI":5,"DIMANCHE":6}
    rows = []
    for _, r in edt_df.iterrows():
        d_idx = py_day.get(str(r[jour_col]).upper(), None)
        if d_idx is None: continue
        # trouve la prochaine occurrence de ce jour
        delta = (d_idx - today_idx) % 7
        day_date = (now + timedelta(days=delta)).date()
        m = time_to_minutes(r["Heure début"])
        if m is None: continue
        dt = datetime.combine(day_date, datetime.min.time()) + timedelta(minutes=m)
        # si séance déjà passée aujourd'hui, pousser à la semaine prochaine
        if dt < now:
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

def subgroup(df, niveau=None, spec=None, groupe=None, semestre="S1"):
    keep = df.copy()
    if semestre: keep = keep[keep["Semestre"].astype(str).str.upper()==semestre.upper()]
    if niveau: keep = keep[keep["Niveau"].astype(str).str.upper()==niveau.upper()]
    if spec:   keep = keep[keep["Spécialité"].astype(str).str.contains(spec, case=False, na=False)]
    if groupe: keep = keep[keep["Groupe"].astype(str).str.upper()==groupe.upper()]
    return keep

# --------------- App ---------------
st.title("🗓️ Portail Génie Civil — EDT & Listes")

edt, etu = load_data()

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)
    st.markdown("---")
    st.caption("Filtres rapides")
    niveau  = st.selectbox("Niveau", sorted(edt["Niveau"].dropna().unique()), index=None, placeholder="Tous")
    spec    = st.selectbox("Spécialité", sorted(edt["Spécialité"].dropna().unique()), index=None, placeholder="Toutes")
    groupe  = st.selectbox("Groupe", sorted(edt["Groupe"].dropna().unique()), index=None, placeholder="Tous")
    st.markdown("---")
    st.caption("Recherche")
    q_nom   = st.text_input("Nom/Prénom (étudiant ou enseignant)")
    st.markdown("---")
    st.caption("Astuce : utilise la loupe de ton navigateur pour filtrer plus vite.")

now = datetime.now()
bloc = subgroup(edt, niveau, spec, groupe, semestre="S1")

# ------------------- Étudiant -------------------
if role == "Étudiant":
    st.header("👩‍🎓 Espace Étudiant")

    col1, col2 = st.columns([2,1])
    with col1:
        st.subheader("Mon EDT")
        st.dataframe(bloc[["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence"]], use_container_width=True)

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
    etu_g = subgroup(etu, niveau, spec, groupe, semestre="S1")[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Aucune liste d’étudiants correspondante.")
    else:
        # feuille de présence locale (session uniquement)
        etu_g["Présent"] = False
        edited = st.data_editor(etu_g, use_container_width=True, height=420, num_rows="fixed")
        st.download_button("⬇️ Télécharger la feuille de présence (CSV)",
                           data=edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"presence_{(niveau or 'X')}_{(groupe or 'X')}_S1.csv",
                           mime="text/csv")
        st.caption("ℹ️ Coche Présent/Absent, puis télécharge le CSV ou fais une capture d’écran.")

    st.subheader("Je cherche mon nom")
    if q_nom:
        hits = etu[etu.apply(lambda r: q_nom.lower() in f"{r['Nom']} {r['Prenom']}".lower(), axis=1)]
        st.dataframe(hits[["Nom","Prenom","Niveau","Spécialité","Groupe","Semestre","Remarque"]], use_container_width=True)

# ------------------- Enseignant -------------------
else:
    st.header("👨‍🏫 Espace Enseignant")

    st.subheader("Planning filtré")
    if q_nom:
        bloc = bloc[bloc["Enseignant"].str.contains(q_nom, case=False, na=False)]
    st.dataframe(bloc[["Jour","Heure début","Heure fin","Matière","Type","Groupe","Salle","Fréquence","Spécialité"]], use_container_width=True)

    st.subheader("Ma prochaine séance")
    nxt = next_session(now, bloc)
    if nxt:
        dt, r = nxt
        st.metric(
            label=f"{r['Jour']} • {r['Heure début']}–{r['Heure fin']}",
            value=f"{r['Matière']} ({r['Type']})",
            delta=f"Dans {human_delta(dt, now)} • Salle {r['Salle']} • Groupe {r['Groupe']}"
        )
    else:
        st.info("Aucune séance trouvée avec les filtres actuels.")

    st.subheader("Trouver une salle pour un enseignant")
    if q_nom:
        salles = edt[edt["Enseignant"].str.contains(q_nom, case=False, na=False)][["Jour","Heure début","Heure fin","Salle","Groupe","Spécialité"]]
        st.dataframe(salles.sort_values(["Jour","Heure début"]), use_container_width=True)

    st.subheader("Liste des étudiants du groupe sélectionné")
    etu_g = subgroup(etu, niveau, spec, groupe, semestre="S1")[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
    if etu_g.empty:
        st.warning("Aucune liste d’étudiants correspondante.")
    else:
        etu_g["Présent"] = False
        edited = st.data_editor(etu_g, use_container_width=True, height=420, num_rows="fixed")
        st.download_button("⬇️ Télécharger la feuille de présence (CSV)",
                           data=edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"presence_{(niveau or 'X')}_{(groupe or 'X')}_S1.csv",
                           mime="text/csv")

st.divider()
st.caption("Données : EDT_MASTER_S1 & ETUDIANTS_MASTER_S1 • Sessions isolées (Streamlit) • Aucune donnée modifiée côté serveur.")
