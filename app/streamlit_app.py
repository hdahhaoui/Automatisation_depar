# =========================
# app/streamlit_app.py
# =========================
# Thème optionnel:
# Crée un fichier .streamlit/config.toml avec:
# [theme]
# base="dark"
# primaryColor="#5eead4"
# backgroundColor="#0b0f17"
# secondaryBackgroundColor="#121826"
# textColor="#e5e7eb"

import re
import glob
from io import BytesIO
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------- Configuration globale ----------------
st.set_page_config(page_title="EDT & Listes • Génie Civil (S1)", page_icon="🗓️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
RAW_EDT = str(BASE_DIR / "data" / "raw" / "edt")
RAW_STU = str(BASE_DIR / "data" / "raw" / "students")
SEMESTRE = "S1"

ORDER_JOUR = {"DIMANCHE":0,"LUNDI":1,"MARDI":2,"MERCREDI":3,"JEUDI":4,"VENDREDI":5,"SAMEDI":6}

# ---------- UI helpers ----------
def inject_css():
    st.markdown("""
    <style>
      h1, h2, h3 { letter-spacing:.2px }
      .actionbar { display:flex; gap:.5rem; flex-wrap:wrap; margin:.25rem 0 1rem }
      .pill { padding:.35rem .6rem; border-radius:999px; background:#1f2937; font-size:.85rem }
      .role-etudiant  { background:#0b3b2e; color:#8ef5dd }
      .role-enseignant{ background:#2a2543; color:#c3b5ff }
      .stDataFrame table { font-size: 0.92rem }
      .card { background:#0f1624; border:1px solid #1f2937; padding:1rem; border-radius:12px }
      .muted{ color:#9ca3af; font-size:.9rem }
      .badge { padding:.2rem .45rem; border-radius:8px; background:#1f2937; font-size:.78rem; margin-left:.35rem }
      .sticky { position:sticky; top:0; z-index:9; background:transparent; padding-top:.25rem }
    </style>
    """, unsafe_allow_html=True)

def header_role(role_label, subtitle):
    role_class = "role-etudiant" if role_label=="Étudiant" else "role-enseignant"
    st.markdown(
        f"""
        <div class="sticky">
          <div class="actionbar">
            <span class="pill {role_class}">{"👩‍🎓 Étudiant" if role_label=="Étudiant" else "👨‍🏫 Enseignant"}</span>
            <span class="pill">{subtitle}</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

inject_css()

# ---------------- Utilitaires ----------------
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
    Infère (Spec2, Niv2, Groupe, Semestre) depuis le nom du fichier.
    Exemple: ETUDIANTS_M1_RIB_G11_S1.xlsx → ("RIB","M1","G11","S1")
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
    if "RIB" in name:    return "RIB", ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "VOA" in name:    return "VOA", ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    if "STRUCT" in name: return "STRUCTURE", ("M2" if "M2" in name else "M1"), g, (sem or "S1")
    # Licence
    if "L2" in name:     return "LICENCE", "2", g, (sem or "S1")
    if "L3" in name:     return "LICENCE", "3", g, (sem or "S1")
    # Ingénieur
    if "1ING" in name:   return "INGENIEUR", "1", g, (sem or "S1")
    if "2ING" in name:   return "INGENIEUR", "2", g, (sem or "S1")
    if "3ING" in name:   return "INGENIEUR", "3", g, (sem or "S1")
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

# ---------------- Chargement des données ----------------
@st.cache_data
def load_raw_s1():
    # ---- EDT
    edt_files = glob.glob(f"{RAW_EDT}/*_S1.*")
    edt_list = []
    for f in edt_files:
        try:
            df = read_any(f)
            df = ensure_cols(df, EDT_COLS, numeric=["Durée (h)"])
            df["Semestre"] = df["Semestre"].apply(normalize_semestre)
            df["Groupe"]   = df["Groupe"].apply(normalize_groupe)
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
            df["Semestre"] = df["Semestre"].apply(normalize_semestre)
            df["Groupe"]   = df["Groupe"].apply(normalize_groupe)
            sp2, lv2 = [], []
            for _, r in df.iterrows():
                s2, n2 = classify_spec_level(r.get("Spécialité",""), r.get("Niveau",""))
                sp2.append(s2); lv2.append(n2)
            df["Spec2"], df["Niv2"] = sp2, lv2
            # fallback via nom de fichier
            if (df["Spec2"] == "").any() or (df["Niv2"] == "").any() or (df["Groupe"] == "").any():
                s2_f, n2_f, g_f, sem_f = infer_from_filename(f)
                if s2_f:  df.loc[df["Spec2"]=="", "Spec2"] = s2_f
                if n2_f:  df.loc[df["Niv2"]=="",  "Niv2"]  = n2_f
                if g_f:   df.loc[df["Groupe"]=="","Groupe"] = normalize_groupe(g_f)
                if sem_f: df.loc[df["Semestre"]=="","Semestre"] = sem_f
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

# ---------------- UI principale ----------------
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

    # Groupes d'abord depuis EDT, sinon fallback listes étudiants
    g_from_edt = subgroup_by_spec_level(edt, spec, niv)["Groupe"].dropna().map(normalize_groupe)
    g_from_etu = subgroup_by_spec_level(etu, spec, niv)["Groupe"].dropna().map(normalize_groupe)
    grp_pool = sorted(pd.concat([g_from_edt, g_from_etu]).unique().tolist())
    groupe = st.selectbox("Groupe", grp_pool, index=0 if grp_pool else None)

    st.markdown("---")
    q_nom = st.text_input("Nom/Prénom (étudiant ou enseignant)")
    st.caption("Astuce : tape un nom puis appuie sur Entrée ⏎")
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

# ---------------- VUE ÉTUDIANT ----------------
if role == "Étudiant":
    # Liste d'étudiants SUPPRIMÉE ici (réservée aux enseignants)
    header_role("Étudiant", f"{title_clean} • Groupe {groupe}")

    # Préparer la vue EDT une fois
    view = bloc[["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Fréquence"]]

    tab_edt, tab_next = st.tabs(["📅 Mon EDT", "⏭️ Prochaine séance"])

    with tab_edt:
        st.markdown("#### Emploi du temps")
        c1, c2 = st.columns([1,1])
        with c1:
            st.download_button("⬇️ Export EDT (CSV)", view.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"EDT_{spec}_{niv}_G{groupe}_S1.csv", use_container_width=True)
        with c2:
            st.download_button("⬇️ Export EDT (XLSX)", df_to_xlsx_bytes(view),
                               file_name=f"EDT_{spec}_{niv}_G{groupe}_S1.xlsx", use_container_width=True)
        st.dataframe(view.rename(columns={"Heure début":"Début","Heure fin":"Fin"}), use_container_width=True, hide_index=True)

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
                    <span class="badge">Enseignant {r['Enseignant']}</span>
                    <span class="badge">Dans {human_delta(dt, now)}</span>
                  </div>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info("Aucune séance à venir avec ces filtres.")

# ---------------- VUE ENSEIGNANT ----------------
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
        c1, c2 = st.columns([1,1])
        with c1:
            st.download_button("⬇️ Export Planning (CSV)",
                               plan_view.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.csv",
                               use_container_width=True)
        with c2:
            st.download_button("⬇️ Export Planning (XLSX)",
                               df_to_xlsx_bytes(plan_view),
                               file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.xlsx",
                               use_container_width=True)
        st.dataframe(plan_view, use_container_width=True, hide_index=True)

    with tab_next:
        st.markdown("#### Ma prochaine séance")
        planning = bloc  # complet pour la prochaine séance du groupe
        nxt = next_session(now, planning)
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
                unsafe_allow_html=True
            )
        else:
            st.info("Aucune séance à venir avec ces filtres.")

    with tab_where:
        st.markdown("#### Recherche d’enseignant par nom")
        if q_nom:
            salles = subgroup_by_spec_level(edt, spec, niv, None)
            salles = salles[salles["Enseignant"].str.contains(q_nom, case=False, na=False)][
                ["Jour","Heure début","Heure fin","Salle","Groupe","Matière"]
            ].sort_values(["Jour","Heure début"])
            st.dataframe(salles, use_container_width=True, hide_index=True)
        else:
            st.info("Saisis un nom dans la barre latérale pour afficher les emplacements.")

    with tab_presence:
        st.markdown("#### Feuille de présence (enseignant)")

        # Charger la liste
        etu_g = subgroup_by_spec_level(etu, spec, niv, groupe)[["N°","Matricule","Nom","Prenom","Remarque"]].reset_index(drop=True)
        if etu_g.empty:
            st.warning("Pas de liste trouvée pour ce groupe (vérifie 'Groupe' = G11/G12 et 'Semestre' = S1).")
        else:
            # Préparer la table dans l'état de session pour pouvoir cocher/décocher tout
            key_df = f"presence_{spec}_{niv}_{groupe}"
            if key_df not in st.session_state:
                df_init = etu_g.copy()
                df_init["Présent"] = False
                st.session_state[key_df] = df_init

            colA, colB, colC = st.columns([1,1,2])
            with colA:
                if st.button("✔️ Tout cocher", use_container_width=True):
                    st.session_state[key_df]["Présent"] = True
            with colB:
                if st.button("✖️ Tout décocher", use_container_width=True):
                    st.session_state[key_df]["Présent"] = False
            with colC:
                st.caption("Astuce : tu peux cocher/décocher ligne par ligne aussi.")

            edited = st.data_editor(
                st.session_state[key_df],
                use_container_width=True,
                height=460,
                num_rows="fixed",
                key=f"editor_{key_df}",
            )
            # Synchroniser l'état
            st.session_state[key_df] = edited

            # Exports
            c1, c2 = st.columns([1,1])
            with c1:
                st.download_button("⬇️ Export présence (CSV)",
                                   edited.to_csv(index=False).encode("utf-8-sig"),
                                   file_name=f"presence_{spec}_{niv}_G{groupe}_S1.csv",
                                   use_container_width=True)
            with c2:
                st.download_button("⬇️ Export présence (XLSX)",
                                   df_to_xlsx_bytes(edited),
                                   file_name=f"presence_{spec}_{niv}_G{groupe}_S1.xlsx",
                                   use_container_width=True)

st.divider()
st.caption("S1 • Spécialité → Niveau → Groupe • Groupes normalisés (G11/G12) • Feuille de présence côté enseignant uniquement.")
