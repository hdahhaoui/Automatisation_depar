# streamlit_app.py  — PART 1/6
import os
import re
import io
import glob
import pytz
import datetime as dt
from typing import Dict, Tuple, Optional

import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# ---------- Config page ----------
st.set_page_config(
    page_title="Portail Génie Civil — EDT & Listes (S1)",
    page_icon="🗓️",
    layout="wide",
)

# ---------- Styles ----------
st.markdown(
    """
    <style>
      .stDownloadButton > button { width: 100%; }
      .compact-badge { 
        display:inline-block; padding:4px 10px; border-radius:999px; 
        background:#e2e8f0; color:#111; font-weight:600; 
        border:1px solid rgba(0,0,0,.05);
      }
      .chip-blue  { background:#4F7BFE; color:white; }
      .chip-green { background:#16a34a; color:white; }
      .chip-purple{ background:#7c3aed; color:white; }
      .muted { color:rgba(0,0,0,.55); }
      .muted-dark { color:rgba(255,255,255,.7); }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------- Répertoires ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EDT_DIR = os.path.join(BASE_DIR, "data", "raw", "edt")
STUD_DIR = os.path.join(BASE_DIR, "data", "raw", "students")

# ---------- Timezone + jours ----------
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

def badge(text, color="chip-blue"):
    return f'<span class="compact-badge {color}">{text}</span>'
# streamlit_app.py  — PART 2/6

# ---------- Normalisation depuis nom fichier ----------
def norm_spec_from_filename(name: str) -> Optional[str]:
    s = re.sub(r"[^a-zA-Z0-9]+", " ", name.upper())
    toks = s.split()
    if "RIB" in toks: return "RIB"
    if "VOA" in toks: return "VOA"
    if "STR" in toks or "STRUCTURE" in toks: return "STRUCTURE"
    if "ING" in toks or "INGENIEUR" in toks: return "INGENIEUR"
    if "L2" in toks or "LICENCE2" in toks or "LICENCE" in toks: return "LICENCE"
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
    if re.search(r"\bG11\b", s): return "G11"
    if re.search(r"\bG12\b", s): return "G12"
    # tolérance si collé
    if "G11" in s: return "G11"
    if "G12" in s: return "G12"
    return None

# ---------- Lecture fichiers ----------
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

        # map colonnes
        colmap = {}
        for c in df.columns:
            cu = str(c).strip().lower()
            if "jour" in cu: colmap[c] = "Jour"
            elif ("début" in cu or "debut" in cu) and "heure" in cu: colmap[c] = "Heure début"
            elif ("fin" in cu) and "heure" in cu: colmap[c] = "Heure fin"
            elif "mati" in cu: colmap[c] = "Matière"
            elif cu in ("type","nature"): colmap[c] = "Type"
            elif "enseign" in cu or "prof" in cu: colmap[c] = "Enseignant"
            elif "salle" in cu or "local" in cu: colmap[c] = "Salle"
            elif "fréq" in cu or "freq" in cu: colmap[c] = "Fréquence"
        df = df.rename(columns=colmap)

        must = ["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle"]
        if not all(m in df.columns for m in must):
            continue

        df["Spécialité"] = spec or ""
        df["Niveau"]     = level or ""
        df["Groupe"]     = group or ""
        df["Jour"]       = df["Jour"].astype(str).str.strip().str.upper()

        keep = must + ["Spécialité","Niveau","Groupe"]
        if "Fréquence" in df.columns: keep += ["Fréquence"]
        rows.append(df[keep])

    if not rows:
        return pd.DataFrame(columns=["Jour","Heure début","Heure fin","Matière","Type","Enseignant","Salle","Spécialité","Niveau","Groupe"])
    return pd.concat(rows, ignore_index=True)

@st.cache_data(show_spinner=False)
def load_all_students() -> Dict[Tuple[str,str,str], pd.DataFrame]:
    out = {}
    for path in sorted(glob.glob(os.path.join(STUD_DIR, "*.xlsx"))):
        fname = os.path.basename(path)
        spec  = norm_spec_from_filename(fname) or "GENIE CIVIL"
        level = norm_level_from_filename(fname) or ""
        group = norm_group_from_filename(fname) or ""

        try:
            df = pd.read_excel(path)
        except Exception:
            continue

        colmap = {}
        for c in df.columns:
            cu = str(c).strip().lower()
            if "nom" in cu and "prénom" in cu: colmap[c] = "NomComplet"
            elif cu.startswith("nom") and "prénom" not in cu: colmap[c] = "Nom"
            elif "prénom" in cu or "prenom" in cu: colmap[c] = "Prénom"
        df = df.rename(columns=colmap)

        if "NomComplet" in df.columns:
            df["Nom affiché"] = df["NomComplet"].astype(str).str.strip()
        else:
            df["Nom affiché"] = (df.get("Nom","").astype(str).str.strip() + " " +
                                 df.get("Prénom","").astype(str).str.strip()).str.strip()
        df = df[["Nom affiché"]]
        df["Nom affiché"] = df["Nom affiché"].replace("nan nan","").str.strip()
        df = df[df["Nom affiché"]!=""].reset_index(drop=True)

        out[(spec, level, group)] = df
    return out

EDT_ALL  = load_all_edt()
STUD_ALL = load_all_students()

# ---------- Hiérarchie filtres ----------
def options_by_hierarchy():
    spec_order = ["INGENIEUR","LICENCE","RIB","VOA","STRUCTURE"]
    specs = [s for s in spec_order if (EDT_ALL["Spécialité"]==s).any() or any(k[0]==s for k in STUD_ALL.keys())]

    levels_by_spec = {}
    for sp in specs:
        levs = set(EDT_ALL.loc[EDT_ALL["Spécialité"]==sp,"Niveau"].unique())
        levs.update([k[1] for k in STUD_ALL.keys() if k[0]==sp])
        levs = [x for x in levs if x]
        order = ["LICENCE 2","LICENCE 3","INGENIEUR 1","INGENIEUR 2","INGENIEUR 3","M1","M2"]
        levs = [l for l in order if l in levs] + [l for l in sorted(levs) if l not in order]
        levels_by_spec[sp] = levs

    groups_by_spec_level = {}
    for sp in specs:
        for lv in levels_by_spec.get(sp,[]):
            gr = set(EDT_ALL.loc[(EDT_ALL["Spécialité"]==sp)&(EDT_ALL["Niveau"]==lv), "Groupe"].unique())
            gr.update([k[2] for k in STUD_ALL.keys() if k[0]==sp and k[1]==lv])
            gr = [g for g in gr if g]
            if "G11" in gr or "G12" in gr:
                order = ["G11","G12"]
                gr = [g for g in order if g in gr] + [g for g in sorted(gr) if g not in order]
            else:
                gr = sorted(gr)
            groups_by_spec_level[(sp,lv)] = gr

    return specs, levels_by_spec, groups_by_spec_level

def filtered_edt(spec, level, group) -> pd.DataFrame:
    df = EDT_ALL.copy()
    if spec:  df = df[df["Spécialité"]==spec]
    if level: df = df[df["Niveau"]==level]
    if group: df = df[df["Groupe"]==group]
    return df.reset_index(drop=True)

def filtered_edt_scope(spec, level, scope: str, group: Optional[str]) -> pd.DataFrame:
    """
    scope: "group" -> juste le groupe ; "all" -> tous les groupes de (spec, level)
    """
    df = EDT_ALL[(EDT_ALL["Spécialité"]==spec) & (EDT_ALL["Niveau"]==level)].copy()
    if scope == "group" and group:
        df = df[df["Groupe"]==group]
    return df.reset_index(drop=True)

def student_list_for(spec, level, group) -> pd.DataFrame:
    return STUD_ALL.get((spec, level, group), pd.DataFrame(columns=["Nom affiché"]))
# streamlit_app.py  — PART 3/6

def filter_area():
    st.sidebar.subheader("🔎 Mode d’accès")
    mode = st.sidebar.radio("Je suis :", ["Étudiant", "Enseignant"], horizontal=False, index=0)

    specs, levels_by_spec, groups_by_spec_level = options_by_hierarchy()

    st.sidebar.markdown("### Filtres hiérarchiques")
    spec = st.sidebar.selectbox("Spécialité", specs, index=0)

    levels = levels_by_spec.get(spec, [])
    if not levels:
        st.sidebar.info("Pas de niveau pour cette spécialité.")
        return mode, spec, None, None, ""
    level = st.sidebar.selectbox("Niveau", levels, index=0, key="sel_level")

    groups = groups_by_spec_level.get((spec, level), [])
    if not groups:
        st.sidebar.info("Pas de groupe pour cette combinaison.")
        return mode, spec, level, None, ""
    group = st.sidebar.selectbox("Groupe", groups, index=0, key="sel_group")

    st.sidebar.markdown("### Nom/Prénom (étudiant ou enseignant)")
    search_name = st.sidebar.text_input(" ", placeholder="tape un nom puis Entrée", label_visibility="collapsed")

    st.sidebar.checkbox("🖨️ Mode impression", value=False, key="print_mode")

    return mode, spec, level, group, search_name

def export_xlsx_bytes(df: pd.DataFrame, sheet_name="Feuille"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf
# streamlit_app.py  — PART 4/6

def render_student_ui(spec, level, group, search_name):
    st.markdown("# Portail Génie Civil — EDT & Listes (S1)")
    st.markdown(
        badge("Étudiant","chip-green") + " " +
        badge(f"{spec} {level} • Groupe {group}","chip-blue"),
        unsafe_allow_html=True
    )

    df = filtered_edt(spec, level, group)

    tabs = st.tabs(["📘 Mon EDT", "⏭️ Prochaine séance"])

    with tabs[0]:
        st.subheader("Emploi du temps")
        if df.empty:
            st.info("Aucun EDT trouvé pour ces filtres.")
        else:
            st.download_button(
                "📥 Exporter l’EDT en Excel",
                data=export_xlsx_bytes(df, "EDT"),
                file_name=f"EDT_{spec}_{level}_{group}_S1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.dataframe(df, use_container_width=True, height=480)

    # Prochaine séance — jour algérien + séance en cours/à venir
    with tabs[1]:
        st.subheader("Prochaine séance (Étudiant)")

        if df.empty:
            st.info("Aucun EDT pour ces filtres.")
            return

        dfe = enrich_times(df)
        default_day = pick_today_label()
        day_choice = st.selectbox("Jour", options=WEEKDAY_FR, index=WEEKDAY_FR.index(default_day))

        day_df = dfe[dfe["Jour"].str.upper()==day_choice].copy()

        def pick_current_and_next(sessions_for_day):
            if sessions_for_day.empty:
                return None, None, "empty"
            tnow = now_dz().time()
            ongoing = sessions_for_day[
                (sessions_for_day["_tstart"] <= tnow) & (tnow < sessions_for_day["_tend"])
            ]
            if not ongoing.empty:
                cur = ongoing.iloc[0]
                nexts = sessions_for_day[sessions_for_day["_tstart"] > cur["_tend"]]
                nxt = nexts.iloc[0] if not nexts.empty else None
                return cur, nxt, "ongoing"
            nxts = sessions_for_day[sessions_for_day["_tstart"] > tnow]
            if not nxts.empty:
                return None, nxts.iloc[0], "upcoming"
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
            st.markdown(f"⏱ **En cours** — reste **{td_to_hm(end_dt - now_local)}** (de {fmt_hhmm(cur['_tstart'])} à {fmt_hhmm(cur['_tend'])}).")
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

def weekly_teacher_view(df_scope: pd.DataFrame, teacher: str) -> pd.DataFrame:
    """Retourne l’EDT hebdomadaire (tous jours) du prof choisi, trié par jour puis heure."""
    if df_scope.empty:
        return df_scope
    dfe = enrich_times(df_scope[df_scope["Enseignant"].astype(str).str.strip().str.lower()==teacher.lower()])
    if dfe.empty:
        return dfe
    # ordonner par jour de la semaine
    dfe["__day_order"] = dfe["Jour"].astype(str).str.upper().apply(lambda j: WEEKDAY_FR.index(j) if j in WEEKDAY_FR else 7)
    dfe = dfe.sort_values(by=["__day_order","_tstart","_tend"]).drop(columns="__day_order")
    return dfe[["Jour","Heure début","Heure fin","Matière","Type","Salle","Groupe"]]

def teachers_today_next(df_day: pd.DataFrame) -> pd.DataFrame:
    """Prochaine (ou en cours) pour chaque enseignant aujourd’hui."""
    if df_day.empty:
        return pd.DataFrame(columns=["Enseignant","Heure début","Heure fin","Matière","Type","Salle","Groupe"])
    dfe = enrich_times(df_day)
    tnow = now_dz().time()
    rows = []
    for teach, grp in dfe.groupby("Enseignant"):
        nxts = grp[grp["_tstart"] > tnow]
        if not nxts.empty:
            r = nxts.iloc[0]
            rows.append(r[["Enseignant","Heure début","Heure fin","Matière","Type","Salle","Groupe"]])
        else:
            ongoing = grp[(grp["_tstart"] <= tnow) & (tnow < grp["_tend"])]
            if not ongoing.empty:
                r = ongoing.iloc[0]
                rows.append(r[["Enseignant","Heure début","Heure fin","Matière","Type","Salle","Groupe"]])
    if not rows:
        return pd.DataFrame(columns=["Enseignant","Heure début","Heure fin","Matière","Type","Salle","Groupe"])
    return pd.DataFrame(rows).sort_values(by=["Enseignant","Heure début"]).reset_index(drop=True)

def render_teacher_ui(spec, level, group, search_name):
    st.markdown("# Portail Génie Civil — EDT & Listes (S1)")
    st.markdown(
        badge("Enseignant","chip-purple") + " " +
        badge(f"{spec} {level} • Groupe {group}","chip-blue"),
        unsafe_allow_html=True
    )

    df_group = filtered_edt(spec, level, group)

    tabs = st.tabs(["🗂️ Planning (ce groupe)", "🧭 Où trouver un enseignant ?", "📝 Feuille de présence"])

    # --- Planning (groupe) ---
    with tabs[0]:
        st.subheader("Planning — Ce groupe")
        if df_group.empty:
            st.info("Aucun EDT pour ces filtres.")
        else:
            st.download_button(
                "📥 Exporter le planning en Excel",
                data=export_xlsx_bytes(df_group, "Planning"),
                file_name=f"Planning_{spec}_{level}_{group}_S1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.dataframe(df_group, use_container_width=True, height=480)

    # --- Où trouver un enseignant ? (liste complète + périmètre) ---
    with tabs[1]:
        st.subheader("Où trouver un enseignant ? (hebdomadaire)")

        perimetre = st.radio(
            "Périmètre",
            ["Ce groupe uniquement", "Tous les groupes de cette spécialité et niveau"],
            index=1, horizontal=True
        )
        scope = "group" if perimetre.startswith("Ce groupe") else "all"
        df_scope = filtered_edt_scope(spec, level, scope, group)

        # Sélecteur enseignant (liste complète du périmètre)
        all_teachers = sorted(t for t in df_scope["Enseignant"].astype(str).unique() if t)
        if search_name:
            s = search_name.strip().lower()
            all_teachers = [t for t in all_teachers if s in t.lower()]
        teacher = st.selectbox("Choisir un enseignant", options=all_teachers, index=0 if all_teachers else None)

        if not all_teachers:
            st.info("Aucun enseignant trouvé pour ce périmètre.")
        else:
            week = weekly_teacher_view(df_scope, teacher)
            if week.empty:
                st.info("Pas de cours pour cet enseignant avec ces filtres.")
            else:
                st.markdown(f"### Planning hebdomadaire — {teacher}")
                st.dataframe(week, use_container_width=True, height=460)

            # Mini vue “aujourd'hui”
            st.markdown("#### Aujourd’hui")
            dftoday = df_scope[df_scope["Jour"].str.upper()==pick_today_label()]
            if not dftoday.empty:
                today_ = weekly_teacher_view(dftoday, teacher)
                if today_.empty:
                    st.info("Pas de séance aujourd’hui pour cet enseignant.")
                else:
                    st.dataframe(today_, use_container_width=True, height=200)
            else:
                st.info("Aucune séance aujourd’hui pour ce périmètre.")

    # --- Feuille de présence ---
    with tabs[2]:
        st.subheader("Feuille de présence (enseignant)")
        stud = student_list_for(spec, level, group)
        if stud.empty:
            st.warning("Aucune liste d’étudiants pour ces filtres.")
            return

        mobile = st.toggle("📱 Mode mobile (affichage compact)", value=True)
        q = st.text_input("🔎 Recherche rapide (Nom/Prénom) :", "")

        key_state  = f"presence::{spec}::{level}::{group}"
        key_remark = f"remarks::{spec}::{level}::{group}"
        if key_state not in st.session_state:
            st.session_state[key_state] = {n: False for n in stud["Nom affiché"].tolist()}
        if key_remark not in st.session_state:
            st.session_state[key_remark] = {n: "" for n in stud["Nom affiché"].tolist()}

        colA, colB = st.columns([1,1])
        with colA:
            if st.button("✅ Tout cocher", use_container_width=True):
                for n in st.session_state[key_state]:
                    st.session_state[key_state][n] = True
        with colB:
            if st.button("❌ Tout décocher", use_container_width=True):
                for n in st.session_state[key_state]:
                    st.session_state[key_state][n] = False

        s = q.strip().lower()
        view = stud.copy()
        if s:
            view = view[view["Nom affiché"].str.lower().str.contains(s)]

        if mobile:
            for name in view["Nom affiché"].tolist():
                cc1, cc2 = st.columns([3,1])
                with cc1:
                    st.write(name)
                with cc2:
                    st.checkbox(
                        "Présent",
                        key=f"chk::{key_state}::{name}",
                        value=st.session_state[key_state].get(name, False),
                        on_change=lambda n=name: st.session_state[key_state].__setitem__(n, not st.session_state[key_state].get(n, False))
                    )
                st.text_input(
                    "Remarque",
                    key=f"rk::{key_remark}::{name}",
                    value=st.session_state[key_remark].get(name,""),
                    on_change=lambda n=name: st.session_state[key_remark].__setitem__(n, st.session_state.get(f"rk::{key_remark}::{n}", "")),
                    label_visibility="collapsed",
                    placeholder="Remarque (facultatif)"
                )
                st.divider()
        else:
            show = []
            for _, r in view.iterrows():
                name = r["Nom affiché"]
                pres = st.session_state[key_state].get(name, False)
                remk = st.session_state[key_remark].get(name, "")
                show.append({"Étudiant": name, "Présent": pres, "Remarque": remk})
            st.dataframe(pd.DataFrame(show), use_container_width=True, height=480)

        # Export PDF (présence)
        def build_presence_pdf() -> bytes:
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            W, H = A4
            margin = 2*cm
            x, y = margin, H - margin

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

            c.setFont("Helvetica-Bold", 10)
            c.drawString(x, y, "N°")
            c.drawString(x+20, y, "Nom & Prénom")
            c.drawString(W - margin - 130, y, "Remarque")
            c.drawString(W - margin - 30, y, "Présent")
            y -= 12
            c.line(x, y, W - margin, y)
            y -= 10

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

def main():
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
