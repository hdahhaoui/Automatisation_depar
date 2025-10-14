# streamlit_app.py  — PART 1/6
import os
import re
import io
import sys
import glob
import time
import math
import uuid
import json
import pytz
import base64
import random
import datetime as dt
from typing import Dict, Tuple, Optional

import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# ---------- Thème & options ----------
st.set_page_config(
    page_title="Portail Génie Civil — EDT & Listes (S1)",
    page_icon="🗓️",
    layout="wide",
)

# Moins de squiggles Streamlit
st.markdown(
    """
    <style>
      .stDownloadButton > button { width: 100%; }
      .compact-badge { 
        display:inline-block; padding:4px 10px; border-radius:999px; 
        background:#e2e8f0; color:#111; font-weight:600; 
        border:1px solid rgba(0,0,0,.05);
      }
      .chip-dark { background:#0ea5e9; color:white; }
      .chip-green { background:#16a34a; color:white; }
      .chip-purple{ background:#7c3aed; color:white; }
      .chip-blue  { background:#4F7BFE; color:white; }
      .chip-orange{ background:#fb923c; color:white; }
      .muted { color:rgba(0,0,0,.55); }
      .muted-dark { color:rgba(255,255,255,.7); }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------- Répertoires de données ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EDT_DIR = os.path.join(BASE_DIR, "data", "raw", "edt")
STUD_DIR = os.path.join(BASE_DIR, "data", "raw", "students")

# ---------- Timezone DZ & jours ----------
TZ_DZ = pytz.timezone("Africa/Algiers")
WEEKDAY_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]

def now_dz() -> dt.datetime:
    return dt.datetime.now(TZ_DZ)

def pick_today_label() -> str:
    return WEEKDAY_FR[now_dz().weekday()]

def parse_hhmm(s: str) -> dt.time:
    if not isinstance(s, str):
        s = str(s or "")
    s = s.strip().lower().replace(" ", "")
    s = s.replace("h", ":")
    if ":" not in s:
        s = f"{s}:00"
    hh, mm = s.split(":", 1)
    return dt.time(int(hh), int(mm or 0))

def td_to_hm(delta: dt.timedelta) -> str:
    secs = max(0, int(delta.total_seconds()))
    h, r = divmod(secs, 3600)
    m, _ = divmod(r, 60)
    return f"{m} min" if h == 0 else f"{h}h{m:02d}"

def fmt_hhmm(t: dt.time) -> str:
    return f"{t.hour:02d}h{t.minute:02d}"

def enrich_times(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    out = df.copy()
    out["_tstart"] = out["Heure début"].apply(parse_hhmm)
    out["_tend"]   = out["Heure fin"].apply(parse_hhmm)
    return out.sort_values(by=["_tstart", "_tend"]).reset_index(drop=True)
# streamlit_app.py  — PART 2/6

# ---------- Normalisation des clés (depuis nom fichiers & colonnes) ----------
SPEC_ALIASES = {
    "GC": "GENIE CIVIL",
    "GENIECIVIL": "GENIE CIVIL",
    "RIB": "RIB",
    "VOA": "VOA",
    "STR": "STRUCTURE",
    "STRUCTURE": "STRUCTURE",
    "LICENCE": "LICENCE",
    "LICENCE2": "LICENCE",
    "LICENCE3": "LICENCE",
    "INGENIEUR": "INGENIEUR",
    "ING": "INGENIEUR",
    "ING1": "INGENIEUR",
    "ING2": "INGENIEUR",
    "ING3": "INGENIEUR",
}

LEVEL_ALIASES = {
    "L2": "LICENCE 2",
    "L3": "LICENCE 3",
    "LICENCE2": "LICENCE 2",
    "LICENCE3": "LICENCE 3",
    "1ING": "INGENIEUR 1",
    "2ING": "INGENIEUR 2",
    "3ING": "INGENIEUR 3",
    "ING1": "INGENIEUR 1",
    "ING2": "INGENIEUR 2",
    "ING3": "INGENIEUR 3",
    "M1": "M1",
    "M2": "M2",
}

def norm_spec_from_filename(name: str) -> Optional[str]:
    s = re.sub(r"[^a-zA-Z0-9]+", " ", name.upper())
    tokens = s.split()
    # ordre de priorité : RIB/VOA/STR → ING → LICENCE
    if "RIB" in tokens: return "RIB"
    if "VOA" in tokens: return "VOA"
    if "STR" in tokens or "STRUCTURE" in tokens: return "STRUCTURE"
    if "ING" in tokens or "INGENIEUR" in tokens: return "INGENIEUR"
    if "L2" in tokens or "LICENCE2" in tokens: return "LICENCE"
    if "L3" in tokens or "LICENCE3" in tokens: return "LICENCE"
    return None

def norm_level_from_filename(name: str) -> Optional[str]:
    s = re.sub(r"[^a-zA-Z0-9]+", " ", name.upper())
    if "L2" in s: return "LICENCE 2"
    if "L3" in s: return "LICENCE 3"
    if "1ING" in s or "ING1" in s: return "INGENIEUR 1"
    if "2ING" in s or "ING2" in s: return "INGENIEUR 2"
    if "3ING" in s or "ING3" in s: return "INGENIEUR 3"
    if "M1" in s: return "M1"
    if "M2" in s: return "M2"
    return None

def norm_group_from_filename(name: str) -> Optional[str]:
    s = name.upper()
    m = re.search(r"\bG(1[12])\b", s)
    if m: return f"G{m.group(1)}"
    # parfois "G11" collé à autre chose :
    m = re.search(r"G11", s)
    if m: return "G11"
    m = re.search(r"G12", s)
    if m: return "G12"
    return None

# ---------- Cache lecture fichiers ----------
@st.cache_data(show_spinner=False)
def load_all_edt() -> pd.DataFrame:
    rows = []
    for path in sorted(glob.glob(os.path.join(EDT_DIR, "*.xlsx"))):
        fname = os.path.basename(path)
        spec  = norm_spec_from_filename(fname) or ""
        level = norm_level_from_filename(fname) or ""
        group = norm_group_from_filename(fname) or ""

        try:
            df = pd.read_excel(path)
        except Exception:
            continue

        # On attend colonnes : Jour, Heure début, Heure fin, Matière, Type, Enseignant, Salle, Fréquence (optionnel)
        colmap = {}
        for c in df.columns:
            cu = str(c).strip().lower()
            if "jour" in cu: colmap[c] = "Jour"
            elif ("début" in cu or "debut" in cu) and "heure" in cu: colmap[c] = "Heure début"
            elif ("fin" in cu) and "heure" in cu: colmap[c] = "Heure fin"
            elif "mati" in cu: colmap[c] = "Matière"
            elif cu in ("type", "nature"): colmap[c] = "Type"
            elif "enseign" in cu or "prof" in cu: colmap[c] = "Enseignant"
            elif "salle" in cu or "local" in cu: colmap[c] = "Salle"
            elif "fréq" in cu or "freq" in cu: colmap[c] = "Fréquence"

        df = df.rename(columns=colmap)
        must = ["Jour", "Heure début", "Heure fin", "Matière", "Type", "Enseignant", "Salle"]
        if not all(m in df.columns for m in must):
            continue

        df["Spécialité"] = spec or ""
        df["Niveau"] = level or ""
        df["Groupe"] = group or ""

        df["Jour"] = df["Jour"].astype(str).str.strip().str.upper()
        rows.append(df[must + ["Spécialité","Niveau","Groupe"] + ([ "Fréquence"] if "Fréquence" in df.columns else [])])

    if not rows:
        return pd.DataFrame(columns=["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Spécialité","Niveau","Groupe"])
    out = pd.concat(rows, ignore_index=True)
    return out

@st.cache_data(show_spinner=False)
def load_all_students() -> Dict[Tuple[str,str,str], pd.DataFrame]:
    """
    Retourne un dict { (spec,level,group): dataframe_etudiants }
    Colonnes attendues : Nom, Prénom (ou Nom & Prénom dans 1 colonne), possibilité Matricule (ignoré pour feuille mobile)
    """
    out: Dict[Tuple[str,str,str], pd.DataFrame] = {}
    for path in sorted(glob.glob(os.path.join(STUD_DIR, "*.xlsx"))):
        fname = os.path.basename(path)
        spec  = norm_spec_from_filename(fname) or "GENIE CIVIL"
        level = norm_level_from_filename(fname) or ""
        group = norm_group_from_filename(fname) or ""

        try:
            df = pd.read_excel(path)
        except Exception:
            continue

        # normaliser colonnes
        colmap = {}
        for c in df.columns:
            cu = str(c).strip().lower()
            if "nom" in cu and "prénom" in cu:  # combine
                colmap[c] = "NomComplet"
            elif cu.startswith("nom") and "prénom" not in cu:
                colmap[c] = "Nom"
            elif "prénom" in cu or "prenom" in cu:
                colmap[c] = "Prénom"
            elif "matricule" in cu or "id" == cu:
                colmap[c] = "Matricule"
        df = df.rename(columns=colmap)

        if "NomComplet" in df.columns:
            # scinder si possible, sinon on garde tel quel
            parts = df["NomComplet"].astype(str).str.strip()
            df["Nom affiché"] = parts
        else:
            df["Nom affiché"] = (df.get("Nom","").astype(str).str.strip() + " " +
                                 df.get("Prénom","").astype(str).str.strip()).str.strip()

        df = df[["Nom affiché"]].copy()
        df["Nom affiché"] = df["Nom affiché"].replace("nan nan","").str.strip()
        df = df[df["Nom affiché"]!=""].reset_index(drop=True)

        key = (spec, level, group)
        out[key] = df

    return out

# Charger au démarrage
EDT_ALL = load_all_edt()
STUD_ALL = load_all_students()

# Options de filtres dynamiques
def options_by_hierarchy():
    # Spécialités
    specs = []
    for s in ["RIB","VOA","STRUCTURE","LICENCE","INGENIEUR"]:
        if (EDT_ALL["Spécialité"]==s).any() or any(k[0]==s for k in STUD_ALL.keys()):
            specs.append(s)

    # Niveaux par spécialité
    levels_by_spec: Dict[str, list] = {}
    for sp in specs:
        levs = set(EDT_ALL.loc[EDT_ALL["Spécialité"]==sp, "Niveau"].unique())
        levs.update([k[1] for k in STUD_ALL.keys() if k[0]==sp])
        levs = [x for x in levs if x]
        order = ["LICENCE 2","LICENCE 3","INGENIEUR 1","INGENIEUR 2","INGENIEUR 3","M1","M2"]
        levs = [l for l in order if l in levs] + [l for l in sorted(levs) if l not in order]
        levels_by_spec[sp] = levs

    # Groupes par (spec, level)
    groups_by_spec_level: Dict[Tuple[str,str], list] = {}
    for sp in specs:
        for lv in levels_by_spec.get(sp, []):
            gr = set(EDT_ALL.loc[(EDT_ALL["Spécialité"]==sp)&(EDT_ALL["Niveau"]==lv), "Groupe"].unique())
            gr.update([k[2] for k in STUD_ALL.keys() if k[0]==sp and k[1]==lv])
            gr = [g for g in gr if g]
            # normaliser : prioriser G11/G12
            if "G11" in gr or "G12" in gr:
                order = ["G11","G12"]
                gr = [g for g in order if g in gr] + [g for g in sorted(gr) if g not in order]
            else:
                gr = sorted(gr)
            groups_by_spec_level[(sp,lv)] = gr

    return specs, levels_by_spec, groups_by_spec_level
# streamlit_app.py  — PART 3/6

def badge(text, color_class="chip-blue"):
    return f'<span class="compact-badge {color_class}">{text}</span>'

def filter_area():
    st.sidebar.subheader("🔎 Mode d’accès")
    mode = st.sidebar.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=False, index=0)

    specs, levels_by_spec, groups_by_spec_level = options_by_hierarchy()

    st.sidebar.markdown("### Filtres hiérarchiques")
    spec = st.sidebar.selectbox("Spécialité", specs, index=0)

    levels = levels_by_spec.get(spec, [])
    if not levels:
        st.sidebar.info("Pas de niveau pour cette spécialité.")
        return mode, spec, None, None
    level = st.sidebar.selectbox("Niveau", levels, index=0, key="sel_level")

    groups = groups_by_spec_level.get((spec, level), [])
    if not groups:
        st.sidebar.info("Pas de groupe pour cette combinaison.")
        return mode, spec, level, None
    group = st.sidebar.selectbox("Groupe", groups, index=0, key="sel_group")

    # Champ de recherche nom (étudiant/enseignant) — optionnel
    st.sidebar.markdown("### Nom/Prénom (étudiant ou enseignant)")
    search_name = st.sidebar.text_input(" ", placeholder="tape un nom puis Entrée", label_visibility="collapsed")

    # Mode impression
    st.sidebar.checkbox("🖨️ Mode impression", value=False, key="print_mode")

    return mode, spec, level, group, search_name

def filtered_edt(spec, level, group) -> pd.DataFrame:
    df = EDT_ALL.copy()
    if spec:  df = df[df["Spécialité"]==spec]
    if level: df = df[df["Niveau"]==level]
    if group: df = df[df["Groupe"]==group]
    return df.reset_index(drop=True)

def student_list_for(spec, level, group) -> pd.DataFrame:
    return STUD_ALL.get((spec, level, group), pd.DataFrame(columns=["Nom affiché"]))
# streamlit_app.py  — PART 4/6

def export_xlsx_bytes(df: pd.DataFrame, sheet_name="EDT"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf

def render_student_ui(spec, level, group, search_name):
    st.markdown(f"# Portail Génie Civil — EDT & Listes (S1)")
    st.markdown(
        badge("Étudiant","chip-green") + " " +
        badge(f"{spec} {level} • Groupe {group}","chip-blue"),
        unsafe_allow_html=True
    )

    df = filtered_edt(spec, level, group)

    tabs = st.tabs(["📘 Mon EDT", "⏭️ Prochaine séance"])
    # --- Mon EDT ---
    with tabs[0]:
        st.subheader("Emploi du temps")
        if df.empty:
            st.info("Aucun EDT trouvé pour ces filtres.")
        else:
            st.download_button(
                "📥 Exporter l’EDT en Excel",
                data=export_xlsx_bytes(df),
                file_name=f"EDT_{spec}_{level}_{group}_S1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.dataframe(df, use_container_width=True, height=480)

    # --- Prochaine séance (Algérie + séance en cours / prochaine / ensuite) ---
    with tabs[1]:
        st.subheader("Prochaine séance (Étudiant)")

        if df.empty:
            st.info("Aucun EDT n'est disponible pour ces filtres.")
            return

        dfe = enrich_times(df)
        all_days = WEEKDAY_FR
        default_day = pick_today_label()
        day_choice = st.selectbox("Jour", options=all_days, index=all_days.index(default_day))

        day_df = dfe[dfe["Jour"].str.upper()==day_choice].copy()
        def pick_current_and_next(sessions_for_day):
            if sessions_for_day.empty:
                return None, None, "empty"
            tnow = now_dz().time()
            ongoing = sessions_for_day[
                (sessions_for_day["_tstart"] <= tnow) &
                (tnow < sessions_for_day["_tend"])
            ]
            if not ongoing.empty:
                cur = ongoing.iloc[0]
                nexts = sessions_for_day[sessions_for_day["_tstart"] > cur["_tend"]]
                nxt = nexts.iloc[0] if not nexts.empty else None
                return cur, nxt, "ongoing"
            nxts = sessions_for_day[sessions_for_day["_tstart"] > tnow]
            if not nxts.empty:
                nxt = nxts.iloc[0]
                return None, nxt, "upcoming"
            return None, None, "completed"

        def block_session(row, color="#4F7BFE"):
            mat = str(row["Matière"])
            typ = str(row.get("Type","") or "")
            teach = str(row.get("Enseignant","") or "")
            salle = str(row.get("Salle","") or "")
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

        cur, nxt, state = pick_current_and_next(day_df)
        st.caption("Prochaine séance")

        now_local = now_dz()
        if state == "empty":
            st.warning(f"Aucune séance planifiée pour {day_choice}.")
        elif state == "ongoing":
            block_session(cur, color="#16a34a")
            end_dt = now_local.replace(hour=cur["_tend"].hour, minute=cur["_tend"].minute, second=0, microsecond=0)
            st.markdown(f"⏱ **En cours** — reste **{td_to_hm(end_dt-now_local)}** (de {fmt_hhmm(cur['_tstart'])} à {fmt_hhmm(cur['_tend'])}).")
            st.markdown("**Après :**")
            if nxt is not None:
                block_session(nxt, color="#4F7BFE")
                start_dt = now_local.replace(hour=nxt["_tstart"].hour, minute=nxt["_tstart"].minute, second=0, microsecond=0)
                st.caption(f"🗓 Dans **{td_to_hm(start_dt - now_local)}** (début {fmt_hhmm(nxt['_tstart'])}).")
            else:
                st.info("Aucune autre séance après celle en cours.")
        elif state == "upcoming":
            block_session(nxt, color="#4F7BFE")
            start_dt = now_local.replace(hour=nxt["_tstart"].hour, minute=nxt["_tstart"].minute, second=0, microsecond=0)
            st.caption(f"🗓 Dans **{td_to_hm(start_dt - now_local)}** (de {fmt_hhmm(nxt['_tstart'])} à {fmt_hhmm(nxt['_tend'])}).")
            rest = day_df[day_df["_tstart"] > nxt["_tstart"]]
            if not rest.empty:
                st.markdown("**Après :**")
                nxt2 = rest.iloc[0]
                block_session(nxt2, color="#7c3aed")
        else:
            st.info(f"Toutes les séances de **{day_choice}** sont terminées.")
# streamlit_app.py  — PART 5/6

def teachers_next_today(df_day: pd.DataFrame) -> pd.DataFrame:
    """Retourne par enseignant la prochaine séance aujourd'hui (après maintenant) + salle."""
    if df_day.empty:
        return pd.DataFrame(columns=["Enseignant","Heure début","Heure fin","Matière","Type","Salle"])
    dfe = enrich_times(df_day)
    tnow = now_dz().time()
    # pour chaque enseignant, choisir la première séance dont _tstart > tnow, sinon la dernière en cours
    rows = []
    for teach, grp in dfe.groupby("Enseignant"):
        nxts = grp[grp["_tstart"] > tnow]
        if not nxts.empty:
            r = nxts.iloc[0]
            rows.append(r[["Enseignant","Heure début","Heure fin","Matière","Type","Salle"]])
        else:
            ongoing = grp[(grp["_tstart"] <= tnow) & (tnow < grp["_tend"])]
            if not ongoing.empty:
                r = ongoing.iloc[0]
                rows.append(r[["Enseignant","Heure début","Heure fin","Matière","Type","Salle"]])
    if not rows:
        return pd.DataFrame(columns=["Enseignant","Heure début","Heure fin","Matière","Type","Salle"])
    out = pd.DataFrame(rows).sort_values(by=["Enseignant","Heure début"]).reset_index(drop=True)
    return out

def render_teacher_ui(spec, level, group, search_name):
    st.markdown(f"# Portail Génie Civil — EDT & Listes (S1)")
    st.markdown(
        badge("Enseignant","chip-purple") + " " +
        badge(f"{spec} {level} • Groupe {group}","chip-blue"),
        unsafe_allow_html=True
    )

    df = filtered_edt(spec, level, group)

    tabs = st.tabs(["🗂️ Planning", "🧭 Où trouver un enseignant ?", "📝 Feuille de présence"])
    # --- Planning ---
    with tabs[0]:
        st.subheader("Planning — Aperçu")
        if df.empty:
            st.info("Aucun EDT pour ces filtres.")
        else:
            st.download_button(
                "📥 Exporter le planning en Excel",
                data=export_xlsx_bytes(df, "Planning"),
                file_name=f"Planning_{spec}_{level}_{group}_S1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.dataframe(df, use_container_width=True, height=480)

    # --- Où trouver un enseignant ? ---
    with tabs[1]:
        st.subheader("Où trouver un enseignant aujourd’hui ?")
        dftoday = df[df["Jour"].str.upper()==pick_today_label()].copy()
        # filtre salle optionnel
        all_salles = sorted(x for x in dftoday["Salle"].astype(str).unique() if x)
        col1, col2 = st.columns([1,1])
        with col1:
            room = st.selectbox("Filtrer par salle (optionnel)", options=["(toutes)"]+all_salles, index=0)
        if room != "(toutes)":
            dftoday = dftoday[dftoday["Salle"].astype(str)==room]
        res = teachers_next_today(dftoday)
        if search_name:
            s = search_name.strip().lower()
            res = res[res["Enseignant"].astype(str).str.lower().str.contains(s)]
        if res.empty:
            st.info("Aucun cours planifié pour aujourd’hui selon ce filtre.")
        else:
            st.dataframe(res, use_container_width=True, height=480)

    # --- Feuille de présence (enseignant) ---
    with tabs[2]:
        st.subheader("Feuille de présence (enseignant)")
        stud = student_list_for(spec, level, group)
        if stud.empty:
            st.warning("Aucune liste d’étudiants pour ces filtres.")
            return

        # Barre outils
        c1, c2, c3 = st.columns([1,2,2])
        with c1:
            mobile = st.toggle("📱 Mode mobile (affichage compact)", value=True, help="Nom + case Présent seulement, + Remarque")
        with c2:
            q = st.text_input("🔎 Recherche rapide (Nom/Prénom) :", "")
        # State présence
        key_state = f"presence::{spec}::{level}::{group}"
        key_remark = f"remarks::{spec}::{level}::{group}"
        if key_state not in st.session_state:
            st.session_state[key_state] = {name: False for name in stud["Nom affiché"].tolist()}
        if key_remark not in st.session_state:
            st.session_state[key_remark] = {name: "" for name in stud["Nom affiché"].tolist()}

        # Actions
        colA, colB = st.columns([1,1])
        with colA:
            if st.button("✅ Tout cocher", use_container_width=True):
                for n in st.session_state[key_state]:
                    st.session_state[key_state][n] = True
        with colB:
            if st.button("❌ Tout décocher", use_container_width=True):
                for n in st.session_state[key_state]:
                    st.session_state[key_state][n] = False

        # Filtrage recherche
        s = q.strip().lower()
        view = stud.copy()
        if s:
            view = view[view["Nom affiché"].str.lower().str.contains(s)]

        # Rendu compact mobile
        if mobile:
            # Liste simple : Nom + case + remarque (une ligne par étudiant)
            data = view["Nom affiché"].tolist()
            for name in data:
                cc1, cc2 = st.columns([3,1])
                with cc1:
                    st.write(name)
                with cc2:
                    st.checkbox("Présent", key=f"chk::{key_state}::{name}", value=st.session_state[key_state].get(name,False),
                                on_change=lambda n=name: st.session_state[key_state].__setitem__(n, not st.session_state[key_state].get(n,False)))
                # remarque
                st.text_input("Remarque", key=f"rk::{key_remark}::{name}",
                              value=st.session_state[key_remark].get(name,""),
                              on_change=lambda n=name: st.session_state[key_remark].__setitem__(n, st.session_state.get(f"rk::{key_remark}::{n}", "")),
                              label_visibility="collapsed",
                              placeholder="Remarque (facultatif)")
                st.divider()
        else:
            # Tableau standard
            show = []
            for _, r in view.iterrows():
                name = r["Nom affiché"]
                pres = st.session_state[key_state].get(name,False)
                remk = st.session_state[key_remark].get(name,"")
                show.append({"Étudiant": name, "Présent": pres, "Remarque": remk})
            st.dataframe(pd.DataFrame(show), use_container_width=True, height=480)

        # Export PDF
        def build_presence_pdf() -> bytes:
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            W, H = A4
            margin = 2*cm
            x, y = margin, H - margin

            # En-tête
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x, y, "UNIVERSITÉ DE TLEMCEN")
            y -= 16
            c.setFont("Helvetica-Bold", 11)
            c.drawString(x, y, "FACULTÉ DE TECHNOLOGIE")
            y -= 14
            c.setFont("Helvetica-Bold", 11)
            c.drawString(x, y, "DÉPARTEMENT DE GÉNIE CIVIL")
            y -= 18
            c.setFont("Helvetica", 10)
            c.drawString(x, y, f"Spécialité : {spec}    Niveau : {level}    Groupe : {group}")
            y -= 14
            dstr = now_dz().strftime("%d/%m/%Y %H:%M")
            c.drawString(x, y, f"Feuille de présence — Date/Heure : {dstr}")
            y -= 10
            c.line(x, y, W - margin, y)
            y -= 18

            # Entêtes de colonnes
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x, y, "N°")
            c.drawString(x+20, y, "Nom & Prénom")
            c.drawString(W - margin - 130, y, "Remarque")
            c.drawString(W - margin - 30, y, "Présent")
            y -= 12
            c.line(x, y, W - margin, y)
            y -= 10

            # Lignes
            c.setFont("Helvetica", 10)
            i = 1
            for name in stud["Nom affiché"].tolist():
                pres = st.session_state[key_state].get(name, False)
                remk = st.session_state[key_remark].get(name, "")
                if y < margin + 40:
                    c.showPage()
                    y = H - margin
                c.drawString(x, y, str(i))
                c.drawString(x+20, y, name[:50])
                c.drawString(W - margin - 130, y, remk[:28])
                c.drawString(W - margin - 30, y, "✓" if pres else "✗")
                y -= 14
                i += 1

            c.showPage()
            c.save()
            buffer.seek(0)
            return buffer.getvalue()

        st.download_button(
            "📄 Exporter la présence en PDF",
            data=build_presence_pdf(),
            file_name=f"Presence_{spec}_{level}_{group}_{now_dz().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
# streamlit_app.py  — PART 6/6

def diagnostic_box():
    with st.expander("Index des fichiers détectés (diagnostic)", expanded=False):
        st.write("Niveaux disponibles par spécialité :")
        specs, levels_by_spec, groups_by_spec_level = options_by_hierarchy()
        st.write(levels_by_spec)

        st.write("Exemples de fichiers EDT :")
        st.write(sorted(os.path.basename(x) for x in glob.glob(os.path.join(EDT_DIR,"*.xlsx"))))

        st.write("Exemples de listes étudiants :")
        st.write(sorted(os.path.basename(x) for x in glob.glob(os.path.join(STUD_DIR,"*.xlsx"))))

def main():
    # diagnostic masqué par défaut
    diagnostic_box()

    mode, spec, level, group, search_name = filter_area()
    if not spec or not level or not group:
        st.info("Choisis une combinaison Spécialité → Niveau → Groupe.")
        return

    if mode == "Étudiant":
        render_student_ui(spec, level, group, search_name)
    else:
        render_teacher_ui(spec, level, group, search_name)

if __name__ == "__main__":
    main()
