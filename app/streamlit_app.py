# ======================================================================
# Portail Génie Civil — EDT & Listes (S1)
# Unique fichier Streamlit : app/streamlit_app.py
# ======================================================================
# Fonctionnalités principales :
# - Profils Étudiant / Enseignant (comportements distincts)
# - Filtres hiérarchiques : Spécialité → Niveau → Groupe
# - Normalisation EDT + Listes étudiants (S1)
# - Harmonisation colonnes listes (Nom/Prénom, Matricule, …)
# - Inférence Spec/Niveau/Groupe/Semestre depuis le nom de fichier
# - Vue Étudiant : Mon EDT, Prochaine séance (+ exports Excel)
# - Vue Enseignant : Planning, Prochaine séance, Où trouver un enseignant ?, Feuille de présence
# - Feuille de présence : case à cocher par étudiant + boutons Tout cocher / Tout décocher
# - Exports uniquement en Excel (.xlsx) (aucun CSV)
# - Option Mode impression
#
# Arborescence attendue :
#   app/
#     streamlit_app.py  ← ce fichier
#     data/
#       raw/
#         edt/        ← fichiers EDT *_S1.xlsx (ou .csv)
#         students/   ← fichiers listes étudiants *_S1.xlsx (ou .csv)
#
# Dépendances (requirements.txt) :
#   streamlit
#   pandas
#   openpyxl
#
# Conseillé : forcer thème sombre via .streamlit/config.toml :
# [theme]
# base = "dark"
# primaryColor = "#5eead4"
# backgroundColor = "#0b0f17"
# secondaryBackgroundColor = "#121826"
# textColor = "#e5e7eb"
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

SEMESTRE = "S1"  # application mono-semestre (S1)

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
          /* data_editor : réduire la densité */
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
    """
    Lecture de fichiers .xlsx ou .csv.
    """
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path)

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """
    DataFrame → bytes .xlsx (openpyxl).
    """
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
    return buf.getvalue()

def time_to_minutes(h: Any) -> Optional[int]:
    """
    '08h30' → 510. Retourne None si invalide.
    """
    s = str(h).strip().lower().replace(" ", "")
    if "h" not in s:
        return None
    try:
        hh, mm = s.split("h")
        return int(hh or 0) * 60 + int(mm or 0)
    except Exception:
        return None

def minutes_to_dt(d: datetime, minutes: int) -> datetime:
    """
    Combine la date d et des minutes depuis minuit → datetime.
    """
    return datetime.combine(d.date(), dtime.min) + timedelta(minutes=minutes)

def human_delta(dt: datetime, now: datetime) -> str:
    """
    Délai lisible '1j 2h 30m'.
    """
    s = int((dt - now).total_seconds())
    d = s // 86400
    s %= 86400
    h = s // 3600
    s %= 3600
    m = s // 60
    parts = []
    if d: parts.append(f"{d}j")
    if h: parts.append(f"{h}h")
    if m: parts.append(f"{m}m")
    return " ".join(parts) or "0m"

def next_session(now: datetime, edt_df: pd.DataFrame) -> Optional[Tuple[datetime, pd.Series]]:
    """
    Prochaine séance dans un DataFrame EDT (hebdomadaire).
    """
    if edt_df.empty:
        return None
    py_day = {"LUNDI": 0, "MARDI": 1, "MERCREDI": 2, "JEUDI": 3, "VENDREDI": 4, "SAMEDI": 5, "DIMANCHE": 6}
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
    """
    Garantit un set minimal de colonnes (crée vides si manquantes).
    """
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
    """
    Détecte (Spécialité, Niveau) à partir des champs.
    """
    S = (spec_text or "").upper()
    L = (level_text or "").upper()
    # Masters
    if "RIB" in S:        return "RIB",        "M1" if "M1" in S + L else ("M2" if "M2" in S + L else "")
    if "VOA" in S:        return "VOA",        "M1" if "M1" in S + L else ("M2" if "M2" in S + L else "")
    if "STRUCT" in S:     return "STRUCTURE",  "M1" if "M1" in S + L else ("M2" if "M2" in S + L else "")
    # Licence
    if "L2" in S + L or "LICENCE 2" in S: return "LICENCE", "2"
    if "L3" in S + L or "LICENCE 3" in S: return "LICENCE", "3"
    # Ingénieur
    if any(k in S + L for k in ["ING", "INGÉ", "INGENIEUR", "INGÉNIEUR"]):
        if "1" in S + L: return "INGENIEUR", "1"
        if "2" in S + L: return "INGENIEUR", "2"
        if "3" in S + L: return "INGENIEUR", "3"
        return "INGENIEUR", ""
    return "", ""

def infer_from_filename(path: str) -> Tuple[Optional[str], Optional[str], Optional[str], str]:
    """
    Infère (Spec2, Niv2, Groupe, Semestre) depuis le nom du fichier :
      - ETUDIANTS_M1_RIB_G11_S1.xlsx → ("RIB","M1","G11","S1")
      - ETUDIANTS_1ING_G12_S1.xlsx   → ("INGENIEUR","1","G12","S1")
      - EDT_L3_G11_S1.xlsx           → ("LICENCE","3","G11","S1")
    """
    name = Path(path).stem.upper().replace("-", "_")
    g = None
    m = re.search(r"_G\s*?(\d+)", name)
    if m: g = f"G{m.group(1)}"
    sem = None
    m = re.search(r"_S\s*?(\d+)", name)
    if m: sem = f"S{m.group(1)}"
    if "RIB" in name:    return "RIB",        ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "VOA" in name:    return "VOA",        ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "STRUCT" in name: return "STRUCTURE",  ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "L2" in name:     return "LICENCE", "2", g, (sem or "S1")
    if "L3" in name:     return "LICENCE", "3", g, (sem or "S1")
    if "1ING" in name:   return "INGENIEUR", "1", g, (sem or "S1")
    if "2ING" in name:   return "INGENIEUR", "2", g, (sem or "S1")
    if "3ING" in name:   return "INGENIEUR", "3", g, (sem or "S1")
    return None, None, g, (sem or "S1")

def level_options_for(spec: str) -> Iterable[str]:
    if spec in ("RIB", "VOA", "STRUCTURE"): return ["M1", "M2"]
    if spec == "LICENCE": return ["2", "3"]
    if spec == "INGENIEUR": return ["1", "2", "3"]
    return []

def pretty_level_label(spec: str, niv: str) -> str:
    if spec == "LICENCE": return f"LICENCE {niv}"
    if spec == "INGENIEUR": return f"INGENIEUR {niv}"
    return niv  # Masters M1/M2

# ------------------ HARMONISATION LISTES ÉTUDIANTS --------------------

def harmonize_student_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renomme les colonnes variées vers le schéma standard et
    scinde 'Nom et Prénom' si besoin.
    """
    mapping: Dict[str, str] = {}
    for c in df.columns:
        k = str(c).strip().lower()
        if k in {"n°", "nº", "n", "num", "numero", "numéro", "n°/ordre"}:
            mapping[c] = "N°"
        elif k in {"matricule", "code", "id", "cne", "apogee", "apogée"}:
            mapping[c] = "Matricule"
        elif k in {"nom"}:
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
        elif k in {"annee", "année", "annee scolaire", "année scolaire"}:
            mapping[c] = "Annee"

    if mapping:
        df = df.rename(columns=mapping)

    # Split Nom & Prenom depuis "NomPrenom" si nécessaire
    if "NomPrenom" in df.columns and (("Nom" not in df.columns) or ("Prenom" not in df.columns)):
        np = df["NomPrenom"].astype(str).str.strip()
        if "Nom" not in df.columns:
            df["Nom"] = np.str.split(r"\s+", n=1, expand=True)[0].fillna("")
        if "Prenom" not in df.columns:
            part = np.str.split(r"\s+", n=1, expand=True)
            df["Prenom"] = (part[1] if part.shape[1] > 1 else "").fillna("")

    # Nettoyage d'espaces
    for col in ["N°", "Matricule", "Nom", "Prenom", "Remarque", "Semestre", "Spécialité", "Niveau", "Groupe"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ------------------ CHARGEMENT ET MISE EN FORME ----------------------

@st.cache_data
def load_raw_s1() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Charge tous les fichiers EDT et étudiants de S1,
    normalise, et renvoie (edt, etu).
    """
    # --------- EDT ---------
    edt_files = glob.glob(f"{RAW_EDT}/*_S1.*")
    edt_list = []
    for f in edt_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
            df["Semestre"] = df["Semestre"].apply(normalize_semestre)
            df["Groupe"] = df["Groupe"].apply(normalize_groupe)
            specs, nivs = zip(*df.apply(lambda r: classify_spec_level(r.get("Spécialité", ""), r.get("Niveau", "")), axis=1))
            df["Spec2"], df["Niv2"] = specs, nivs

            # fallback via fichier si besoin
            if (df["Spec2"] == "").any() or (df["Niv2"] == "").any() or (df["Groupe"] == "").any():
                s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
                if s2_f: df.loc[df["Spec2"] == "", "Spec2"] = s2_f
                if n2_f: df.loc[df["Niv2"] == "", "Niv2"] = n2_f
                if g_f:  df.loc[df["Groupe"] == "", "Groupe"] = normalize_groupe(g_f)
                if sem_f: df.loc[df["Semestre"] == "", "Semestre"] = sem_f

            edt_list.append(df)
        except Exception as e:
            st.warning(f"EDT ignoré: {Path(f).name} ({e})")

    if edt_list:
        edt = pd.concat(edt_list, ignore_index=True)
        edt["__o"] = edt["Jour"].map(ORDER_JOUR).fillna(99)
        edt = edt.sort_values(["Spec2", "Niv2", "Groupe", "__o", "Heure début"]).drop(columns="__o")
    else:
        edt = pd.DataFrame(columns=EDT_COLS + ["Spec2", "Niv2"])

    # --------- Étudiants ---------
    stu_files = glob.glob(f"{RAW_STU}/*_S1.*")
    stu_list = []
    for f in stu_files:
        try:
            df = read_any(f)
            df = harmonize_student_columns(df)
            df = ensure_cols(df, STU_COLS)
            df["Semestre"] = df["Semestre"].apply(normalize_semestre)
            df["Groupe"] = df["Groupe"].apply(normalize_groupe)
            specs, nivs = zip(*df.apply(lambda r: classify_spec_level(r.get("Spécialité", ""), r.get("Niveau", "")), axis=1))
            df["Spec2"], df["Niv2"] = specs, nivs

            # fallback fichier
            if (df["Spec2"] == "").any() or (df["Niv2"] == "").any() or (df["Groupe"] == "").any():
                s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
                if s2_f: df.loc[df["Spec2"] == "", "Spec2"] = s2_f
                if n2_f: df.loc[df["Niv2"] == "", "Niv2"] = n2_f
                if g_f:  df.loc[df["Groupe"] == "", "Groupe"] = normalize_groupe(g_f)
                if sem_f: df.loc[df["Semestre"] == "", "Semestre"] = sem_f

            stu_list.append(df)
        except Exception as e:
            st.warning(f"Liste ignorée: {Path(f).name} ({e})")

    if stu_list:
        etu = pd.concat(stu_list, ignore_index=True)
    else:
        etu = pd.DataFrame(columns=STU_COLS + ["Spec2", "Niv2"])

    return edt, etu

def subgroup_by_spec_level(df: pd.DataFrame, spec: Optional[str] = None,
                           niv: Optional[str] = None, groupe: Optional[str] = None) -> pd.DataFrame:
    keep = df[df["Semestre"].astype(str).str.upper() == SEMESTRE]
    if spec:
        keep = keep[keep["Spec2"] == spec]
    if niv:
        keep = keep[keep["Niv2"] == niv]
    if groupe:
        gnorm = normalize_groupe(groupe)
        keep = keep[keep["Groupe"].apply(normalize_groupe) == gnorm]
    return keep

# ----------------------------- UI GLOBALE -----------------------------

st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")

edt, etu = load_raw_s1()

if edt.empty:
    st.error("Aucun EDT S1 trouvé dans `app/data/raw/edt/`.")
    st.stop()

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role = st.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=True)

    st.markdown("---")
    st.caption("Filtres hiérarchiques")

    spec_order = ["RIB", "VOA", "STRUCTURE", "LICENCE", "INGENIEUR"]
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

    # ---- Mon EDT
    with tab_edt:
        st.markdown("#### Emploi du temps")
        view = bloc[["Jour", "Heure début", "Heure fin", "Matière", "Type", "Enseignant", "Salle", "Fréquence"]]
        st.download_button(
            "⬇️ Exporter l’EDT en Excel",
            df_to_xlsx_bytes(view),
            file_name=f"EDT_{spec}_{niv}_G{groupe}_S1.xlsx",
            use_container_width=True,
        )
        st.dataframe(view.rename(columns={"Heure début": "Début", "Heure fin": "Fin"}), use_container_width=True, hide_index=True)

    # ---- Prochaine séance
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

    # ---- Planning
    with tab_plan:
        st.markdown("#### Planning filtré")
        planning = bloc.copy()
        if q_nom:
            planning = planning[planning["Enseignant"].str.contains(q_nom, case=False, na=False)]
        plan_view = planning[["Jour", "Heure début", "Heure fin", "Matière", "Type", "Salle", "Groupe"]]
        st.download_button(
            "⬇️ Exporter le planning en Excel",
            df_to_xlsx_bytes(plan_view),
            file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.xlsx",
            use_container_width=True,
        )
        st.dataframe(plan_view, use_container_width=True, hide_index=True)

    # ---- Prochaine séance (enseignant, côté groupe sélectionné)
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

    # ---- Où trouver un enseignant ? (liste complète du département S1)
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
            base["__end"] = base["Heure fin"].map(time_to_minutes)

            today_idx = now.weekday()
            today_name = JOURS_FR[today_idx]
            now_min = now.hour * 60 + now.minute

            def next_occurrence(row: pd.Series) -> Tuple[Optional[datetime], bool, bool]:
                """
                Retourne (dt_start, is_today, is_now) pour la prochaine occurrence de ce créneau.
                """
                d_idx = ORDER_JOUR.get(str(row["Jour"]).upper(), None)
                if d_idx is None or pd.isna(row["__start"]) or pd.isna(row["__end"]):
                    return None, False, False
                is_today = (d_idx == today_idx)

                # en cours ?
                if is_today and (row["__start"] <= now_min < row["__end"]):
                    dt_start = datetime.combine(now.date(), dtime.min) + timedelta(minutes=int(row["__start"]))
                    return dt_start, True, True

                # prochain
                delta = (d_idx - today_idx) % 7
                dt_day = (now + timedelta(days=delta)).date()
                dt_start = datetime.combine(dt_day, dtime.min) + timedelta(minutes=int(row["__start"]))
                if is_today and row["__start"] < now_min:
                    dt_start += timedelta(days=7)
                    is_today = False
                return dt_start, is_today, False

            rows = []
            for ens, g in base.groupby("Enseignant", dropna=True):
                if ens is None or str(ens).strip() == "":
                    continue

                best = None
                for _, r in g.iterrows():
                    dt, is_today, is_now = next_occurrence(r)
                    if dt is None:
                        continue
                    cand = (dt, is_today, is_now, r)
                    if best is None or cand[0] < best[0]:
                        best = cand

                if best is None:
                    continue

                dt, is_today, is_now, r = best
                if is_now:
                    statut = f"🟢 En cours jusqu’à {r['Heure fin']} (Salle {r['Salle']})"
                    order = (0, dt)
                    jour_txt = r["Jour"].title()
                    heure_txt = f"{r['Heure début']}–{r['Heure fin']}"
                else:
                    if is_today:
                        statut = f"🔵 Dans {human_delta(dt, now)}"
                        jour_txt = "Aujourd’hui"
                    else:
                        statut = f"⚪ Le {r['Jour'].title()} à {r['Heure début']} (dans {human_delta(dt, now)})"
                        jour_txt = r["Jour"].title()
                    order = (1 if is_today else 2, dt)
                    heure_txt = f"{r['Heure début']}–{r['Heure fin']}"

                rows.append({
                    "Enseignant": ens,
                    "Statut": statut,
                    "Heure": heure_txt,
                    "Jour": jour_txt,
                    "Salle": r["Salle"],
                    "Matière": r["Matière"],
                    "Groupe": r["Groupe"],
                    "Spécialité": r.get("Spec2", ""),
                    "Niveau": pretty_level_label(r.get("Spec2", ""), r.get("Niv2", "")),
                    "_order0": order[0],
                    "_order1": order[1],
                })

            df_where = pd.DataFrame(rows)

            if only_today and not df_where.empty:
                df_where = df_where[df_where["Jour"].isin(["Aujourd’hui", today_name.title()])]

            if q_nom and not df_where.empty:
                df_where = df_where[df_where["Enseignant"].str.contains(q_nom, case=False, na=False)]

            if df_where.empty:
                msg_day = "aujourd’hui" if only_today else today_name.title()
                st.info(f"Aucun enseignant à afficher pour **{msg_day}** avec les filtres actuels.")
            else:
                df_where = df_where.sort_values(by=["_order0", "_order1", "Enseignant"])
                st.dataframe(
                    df_where.drop(columns=["_order0", "_order1"]),
                    use_container_width=True,
                    hide_index=True
                )

    # ---- Feuille de présence (enseignant)
    with tab_presence:
        st.markdown("#### Feuille de présence (enseignant)")

        etu_g_raw = subgroup_by_spec_level(etu, spec, niv, groupe).copy()

        base_cols = [c for c in ["N°", "Matricule", "Nom", "Prenom", "Remarque"] if c in etu_g_raw.columns]
        if not base_cols:
            # fallback minimal (afficher au moins quelque chose)
            base_cols = [c for c in etu_g_raw.columns if c not in {"Spec2", "Niv2", "Semestre"}][:5]

        etu_g = etu_g_raw[base_cols].reset_index(drop=True)

        if etu_g.empty:
            st.warning("Pas de liste trouvée pour ce groupe (vérifie 'Groupe' = G11/G12 et 'Semestre' = S1).")
        else:
            key_df = f"presence_{spec}_{niv}_{groupe}"
            if key_df not in st.session_state:
                df_init = etu_g.copy()
                if "Présent" not in df_init.columns:
                    df_init["Présent"] = False
                st.session_state[key_df] = df_init

            colA, colB, colC = st.columns([1, 1, 2])
            with colA:
                if st.button("✔️ Tout cocher", use_container_width=True):
                    st.session_state[key_df]["Présent"] = True
            with colB:
                if st.button("✖️ Tout décocher", use_container_width=True):
                    st.session_state[key_df]["Présent"] = False
            with colC:
                st.caption("Astuce : tu peux cocher/décocher ligne par ligne.")

            edited = st.data_editor(
                st.session_state[key_df],
                use_container_width=True,
                height=460,
                num_rows="fixed",
                key=f"editor_{key_df}",
            )
            st.session_state[key_df] = edited

            st.download_button(
                "⬇️ Exporter la présence en Excel",
                df_to_xlsx_bytes(edited),
                file_name=f"Presence_{spec}_{niv}_G{groupe}_S1.xlsx",
                use_container_width=True,
            )

# ----------------------------- FOOTER ---------------------------------

st.divider()
st.caption(
    "S1 • Spécialité → Niveau → Groupe • Groupes normalisés (G11/G12) • "
    "Harmonisation des listes étudiants • Exports uniquement en Excel (.xlsx) • "
    "Feuille de présence côté enseignant uniquement."
)
