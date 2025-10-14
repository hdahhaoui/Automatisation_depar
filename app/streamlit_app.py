# streamlit_app.py — PART 1/6

import os
import re
import io
import glob
import pytz
import datetime as dt
from typing import Dict, Tuple, Optional, List

import pandas as pd
import streamlit as st

# PDF presence
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# ---------------- Page config
st.set_page_config(
    page_title="Portail Génie Civil — EDT & Listes (S1)",
    page_icon="🗓️",
    layout="wide",
)

# ---------------- Styles
st.markdown(
    """
    <style>
      .stDownloadButton > button { width: 100%; }
      .chip { display:inline-block; padding:4px 10px; border-radius:999px;
              font-weight:600; border:1px solid rgba(0,0,0,.06); }
      .chip-blue  { background:#4F7BFE; color:white; }
      .chip-green { background:#16a34a; color:white; }
      .chip-purple{ background:#7c3aed; color:white; }
      .muted { color:rgba(0,0,0,.55); }
      .linecard { background:#4F7BFE; color:#fff; border-radius:10px; padding:10px 14px; margin-top:8px; }
      .linecard-green { background:#16a34a; }
      .pill { display:inline-block; padding:2px 8px; border-radius:8px; background:#eef2ff; color:#111; }
    </style>
    """,
    unsafe_allow_html=True
)

def chip(txt, cls="chip-blue"):
    return f'<span class="chip {cls}">{txt}</span>'

# ---------------- Dossiers
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EDT_DIR  = os.path.join(BASE_DIR, "data", "raw", "edt")
STUD_DIR = os.path.join(BASE_DIR, "data", "raw", "students")

# ---------------- Temps (Algérie)
TZ_DZ = pytz.timezone("Africa/Algiers")
WEEKDAY_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]

def now_dz() -> dt.datetime:
    return dt.datetime.now(TZ_DZ)

def pick_today_label() -> str:
    return WEEKDAY_FR[now_dz().weekday()]

def parse_hhmm(val) -> dt.time:
    s = str(val or "").strip().lower().replace(" ", "")
    if not s:
        return dt.time(0,0)
    s = s.replace("h", ":")
    if ":" not in s:
        s = f"{s}:00"
    hh, mm = s.split(":", 1)
    return dt.time(int(hh), int(mm or 0))

def fmt_hhmm(t: dt.time) -> str:
    return f"{t.hour:02d}h{t.minute:02d}"

def td_to_hm(delta: dt.timedelta) -> str:
    secs = max(0, int(delta.total_seconds()))
    h, r = divmod(secs, 3600)
    m, _ = divmod(r, 60)
    return f"{m} min" if h == 0 else f"{h}h{m:02d}"

def enrich_times(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df.copy()
    out = df.copy()
    out["_tstart"] = out["Heure début"].apply(parse_hhmm)
    out["_tend"]   = out["Heure fin"].apply(parse_hhmm)
    return out.sort_values(by=["_tstart","_tend"]).reset_index(drop=True)

MIN_COLS = ["Jour","Heure début","Heure fin","Matière","Enseignant","Salle"]
# streamlit_app.py — PART 2/6

def norm_spec_from_filename(name: str) -> Optional[str]:
    s = re.sub(r"[^a-zA-Z0-9]+", " ", name.upper())
    toks = s.split()
    if "RIB" in toks: return "RIB"
    if "VOA" in toks: return "VOA"
    if "STR" in toks or "STRUCTURE" in toks: return "STRUCTURE"
    # forcer INGENIEUR même si forme "ING" / "INGENIEUR"
    if "INGENIEUR" in toks or "ING" in toks: return "INGENIEUR"
    if "LICENCE" in toks or "L2" in toks or "L3" in toks: return "LICENCE"
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
    if "G11" in s: return "G11"
    if "G12" in s: return "G12"
    return None

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
        if not all(m in df.columns for m in must):  # fichier incomplet
            continue

        df["Spécialité"] = spec
        df["Niveau"]     = level
        df["Groupe"]     = group
        df["Jour"]       = df["Jour"].astype(str).str.strip().str.upper()

        keep = ["Jour","Heure début","Heure fin","Matière","Enseignant","Salle","Type","Spécialité","Niveau","Groupe"]
        rows.append(df[keep])

    if not rows:
        return pd.DataFrame(columns=keep)
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
            df["Nom affiché"] = (df.get("Nom","").astype(str).str.strip()+" "+df.get("Prénom","").astype(str).str.strip()).str.strip()
        df = df[["Nom affiché"]]
        df["Nom affiché"] = df["Nom affiché"].replace("nan nan","").str.strip()
        df = df[df["Nom affiché"]!=""].reset_index(drop=True)

        out[(spec, level, group)] = df
    return out

EDT_ALL  = load_all_edt()
STUD_ALL = load_all_students()

def options_by_hierarchy():
    # ordre logique
    spec_order = ["INGENIEUR","LICENCE","RIB","VOA","STRUCTURE"]
    specs = []
    for s in spec_order:
        present = (EDT_ALL["Spécialité"]==s).any() or any(k[0]==s for k in STUD_ALL.keys())
        if present: specs.append(s)

    levels_by_spec = {}
    for sp in specs:
        levs = set(EDT_ALL.loc[EDT_ALL["Spécialité"]==sp,"Niveau"].unique())
        levs.update([k[1] for k in STUD_ALL if k[0]==sp])
        levs = [l for l in levs if l]
        order = ["LICENCE 2","LICENCE 3","INGENIEUR 1","INGENIEUR 2","INGENIEUR 3","M1","M2"]
        levs_sorted = [l for l in order if l in levs] + [l for l in sorted(levs) if l not in order]
        levels_by_spec[sp] = levs_sorted

    groups_by_spec_level = {}
    for sp in specs:
        for lv in levels_by_spec.get(sp,[]):
            gr = set(EDT_ALL.loc[(EDT_ALL["Spécialité"]==sp)&(EDT_ALL["Niveau"]==lv),"Groupe"].unique())
            gr.update([k[2] for k in STUD_ALL if k[0]==sp and k[1]==lv])
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
    df = EDT_ALL[(EDT_ALL["Spécialité"]==spec)&(EDT_ALL["Niveau"]==level)].copy()
    if scope == "group" and group:
        df = df[df["Groupe"]==group]
    return df.reset_index(drop=True)

def student_list_for(spec, level, group) -> pd.DataFrame:
    return STUD_ALL.get((spec, level, group), pd.DataFrame(columns=["Nom affiché"]))
# streamlit_app.py — PART 3/6

def filter_area():
    st.sidebar.subheader("🔎 Mode d’accès")
    mode = st.sidebar.radio("Je suis :", ["Étudiant", "Enseignant"], index=0)

    specs, levels_by_spec, groups_by_spec_level = options_by_hierarchy()
    st.sidebar.markdown("### Filtres hiérarchiques")
    if not specs:
        st.sidebar.error("Aucune spécialité détectée.")
        return mode, None, None, None, ""
    spec = st.sidebar.selectbox("Spécialité", specs, index=0)

    levs = levels_by_spec.get(spec, [])
    if not levs:
        st.sidebar.info("Pas de niveau pour cette spécialité.")
        return mode, spec, None, None, ""
    level = st.sidebar.selectbox("Niveau", levs, index=0)

    groups = groups_by_spec_level.get((spec, level), [])
    if not groups:
        st.sidebar.info("Pas de groupe pour cette combinaison.")
        return mode, spec, level, None, ""
    group = st.sidebar.selectbox("Groupe", groups, index=0)

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
# streamlit_app.py — PART 4/6

def df_minimal(df: pd.DataFrame) -> pd.DataFrame:
    # colonnes utiles uniquement
    keep = [c for c in MIN_COLS if c in df.columns]
    out = df[keep].copy()
    return out

def card_session(row, day_label, color="#4F7BFE"):
    t1, t2 = row["_tstart"], row["_tend"]
    st.markdown(
        f"""
        <div class="linecard" style="background:{color}">
          <strong>{row['Matière']}</strong><span class="pill"> {row.get('Type','Cours')} </span><br/>
          👨‍🏫 {row.get('Enseignant','')} &nbsp; • &nbsp; 🏫 Salle {row.get('Salle','')} &nbsp; • &nbsp; 📅 {day_label}<br/>
          🕒 {fmt_hhmm(t1)} – {fmt_hhmm(t2)}
        </div>
        """, unsafe_allow_html=True
    )

def render_student_ui(spec, level, group, search_name):
    st.markdown("# Portail Génie Civil — EDT & Listes (S1)")
    st.markdown(
        chip("Étudiant","chip-green")+" "+
        chip(f"{spec} {level} • Groupe {group}","chip-blue"),
        unsafe_allow_html=True
    )

    df = filtered_edt(spec, level, group)
    tabs = st.tabs(["📘 Mon EDT", "⏭️ Prochaine séance"])

    with tabs[0]:
        st.subheader("Mon EDT")
        if df.empty:
            st.info("Aucun EDT trouvé.")
        else:
            dfv = df_minimal(df)
            st.download_button(
                "📥 Exporter l’EDT (Excel)",
                data=export_xlsx_bytes(dfv, "EDT"),
                file_name=f"EDT_{spec}_{level}_{group}_S1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.dataframe(dfv, use_container_width=True, height=480)

    with tabs[1]:
        st.subheader("Prochaine séance (Étudiant)")
        if df.empty:
            st.info("Aucun EDT pour ces filtres.")
            return
        dfe = enrich_times(df)
        the_day = st.selectbox("Jour", WEEKDAY_FR, index=WEEKDAY_FR.index(pick_today_label()))
        day_df = dfe[dfe["Jour"].str.upper()==the_day].copy()

        def pick_current_next(daydf):
            if daydf.empty: return None, None, "empty"
            tnow = now_dz().time()
            ongoing = daydf[(daydf["_tstart"] <= tnow) & (tnow < daydf["_tend"])]
            if not ongoing.empty:
                cur = ongoing.iloc[0]
                nxts = daydf[daydf["_tstart"] > cur["_tend"]]
                nxt = nxts.iloc[0] if not nxts.empty else None
                return cur, nxt, "ongoing"
            nxts = daydf[daydf["_tstart"] > tnow]
            if not nxts.empty: return None, nxts.iloc[0], "upcoming"
            return None, None, "done"

        cur, nxt, state = pick_current_next(day_df)
        now_local = now_dz()

        if state == "empty":
            st.warning(f"Aucune séance planifiée pour {the_day}.")
        elif state == "ongoing":
            card_session(cur, the_day, color="#16a34a")
            end_dt = now_local.replace(hour=cur["_tend"].hour, minute=cur["_tend"].minute, second=0, microsecond=0)
            st.caption(f"⏱ En cours — reste **{td_to_hm(end_dt - now_local)}** (de {fmt_hhmm(cur['_tstart'])} à {fmt_hhmm(cur['_tend'])}).")
            if nxt is not None:
                st.markdown("**Après :**")
                card_session(nxt, the_day, color="#4F7BFE")
        elif state == "upcoming":
            card_session(nxt, the_day, color="#4F7BFE")
            start_dt = now_local.replace(hour=nxt["_tstart"].hour, minute=nxt["_tstart"].minute, second=0, microsecond=0)
            st.caption(f"🗓 Dans **{td_to_hm(start_dt - now_local)}** (de {fmt_hhmm(nxt['_tstart'])} à {fmt_hhmm(nxt['_tend'])}).")
        else:
            st.info(f"Toutes les séances de **{the_day}** sont terminées.")
# streamlit_app.py — PART 5/6

def weekly_teacher_view(df_scope: pd.DataFrame, teacher: str) -> pd.DataFrame:
    if df_scope.empty: return df_scope
    filt = df_scope[df_scope["Enseignant"].astype(str).str.strip().str.lower()==teacher.lower()]
    if filt.empty: return filt
    dfe = enrich_times(filt)
    dfe["__d"] = dfe["Jour"].astype(str).str.upper().apply(lambda j: WEEKDAY_FR.index(j) if j in WEEKDAY_FR else 7)
    dfe = dfe.sort_values(by=["__d","_tstart","_tend"]).drop(columns="__d")
    return df_minimal(dfe)

def teachers_today_next(df_day: pd.DataFrame) -> pd.DataFrame:
    if df_day.empty:
        return pd.DataFrame(columns=MIN_COLS)
    dfe = enrich_times(df_day)
    tnow = now_dz().time()
    kept: List[pd.Series] = []
    for teach, grp in dfe.groupby("Enseignant"):
        nxts = grp[grp["_tstart"] > tnow]
        if not nxts.empty: kept.append(nxts.iloc[0])
        else:
            ong = grp[(grp["_tstart"] <= tnow) & (tnow < grp["_tend"])]
            if not ong.empty: kept.append(ong.iloc[0])
    if not kept:
        return pd.DataFrame(columns=MIN_COLS)
    dfk = pd.DataFrame(kept)
    return df_minimal(dfk.sort_values(by=["Enseignant","Heure début"]))

def render_teacher_ui(spec, level, group, search_name):
    st.markdown("# Portail Génie Civil — EDT & Listes (S1)")
    st.markdown(
        chip("Enseignant","chip-purple")+" "+
        chip(f"{spec} {level} • Groupe {group}","chip-blue"),
        unsafe_allow_html=True
    )

    df_group = filtered_edt(spec, level, group)
    tabs = st.tabs(["🗂️ Planning (ce groupe)", "🧭 Où trouver un enseignant ?", "📝 Feuille de présence"])

    # --- Planning groupe (colonnes minimales)
    with tabs[0]:
        st.subheader("Planning — Ce groupe")
        if df_group.empty:
            st.info("Aucun EDT pour ces filtres.")
        else:
            dfv = df_minimal(df_group)
            st.download_button(
                "📥 Exporter le planning (Excel)",
                data=export_xlsx_bytes(dfv, "Planning"),
                file_name=f"Planning_{spec}_{level}_{group}_S1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.dataframe(dfv, use_container_width=True, height=480)

    # --- Où trouver un enseignant ? (liste complète + périmètre)
    with tabs[1]:
        st.subheader("Où trouver un enseignant ? (hebdomadaire)")
        perimeter = st.radio(
            "Périmètre",
            ["Ce groupe uniquement", "Tous les groupes de cette spécialité et niveau"],
            index=1, horizontal=True
        )
        scope = "group" if perimeter.startswith("Ce groupe") else "all"
        df_scope = filtered_edt_scope(spec, level, scope, group)

        teachers = sorted(t for t in df_scope["Enseignant"].astype(str).unique() if t)
        if search_name:
            s = search_name.strip().lower()
            teachers = [t for t in teachers if s in t.lower()]

        teacher = st.selectbox("Choisir un enseignant", options=teachers, index=0 if teachers else None)
        if not teachers:
            st.info("Aucun enseignant trouvé pour ce périmètre.")
        else:
            week = weekly_teacher_view(df_scope, teacher)
            if week.empty:
                st.info("Pas de cours pour cet enseignant.")
            else:
                st.markdown(f"### Planning hebdomadaire — {teacher}")
                st.dataframe(week, use_container_width=True, height=460)

            st.markdown("#### Aujourd’hui")
            dftoday = df_scope[df_scope["Jour"].str.upper()==pick_today_label()]
            tday = weekly_teacher_view(dftoday, teacher)
            if tday.empty:
                st.info("Pas de séance aujourd’hui pour cet enseignant.")
            else:
                st.dataframe(tday, use_container_width=True, height=200)

    # --- Feuille de présence (mobile-friendly + Tout cocher OK + PDF)
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

        def set_all(val: bool):
            for n in stud["Nom affiché"].tolist():
                st.session_state[key_state][n] = val
            st.rerun()

        c1, c2 = st.columns(2)
        with c1:
            st.button("✅ Tout cocher", use_container_width=True, on_click=set_all, args=(True,))
        with c2:
            st.button("❌ Tout décocher", use_container_width=True, on_click=set_all, args=(False,))

        # filtrage nom
        s = q.strip().lower()
        view = stud.copy()
        if s:
            view = view[view["Nom affiché"].str.lower().str.contains(s)]

        if mobile:
            for name in view["Nom affiché"].tolist():
                r1, r2 = st.columns([3,1])
                with r1:
                    st.write(name)
                with r2:
                    st.session_state[key_state][name] = st.checkbox(
                        "Présent", key=f"chk::{key_state}::{name}",
                        value=st.session_state[key_state].get(name, False)
                    )
                st.session_state[key_remark][name] = st.text_input(
                    "Remarque", key=f"rk::{key_remark}::{name}",
                    value=st.session_state[key_remark].get(name, ""),
                    label_visibility="collapsed",
                    placeholder="Remarque (facultatif)"
                )
                st.divider()
        else:
            data = []
            for name in view["Nom affiché"].tolist():
                data.append({
                    "Étudiant": name,
                    "Présent": st.session_state[key_state].get(name, False),
                    "Remarque": st.session_state[key_remark].get(name, "")
                })
            st.dataframe(pd.DataFrame(data), use_container_width=True, height=480)

        # ----- Export PDF (en-tête Université / Fac / Département)
        def build_presence_pdf() -> bytes:
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            W, H = A4
            margin = 2*cm
            x, y = margin, H - margin

            c.setFont("Helvetica-Bold", 12); c.drawString(x, y, "UNIVERSITÉ DE TLEMCEN"); y -= 16
            c.setFont("Helvetica-Bold", 11); c.drawString(x, y, "FACULTÉ DE TECHNOLOGIE"); y -= 14
            c.setFont("Helvetica-Bold", 11); c.drawString(x, y, "DÉPARTEMENT DE GÉNIE CIVIL"); y -= 18
            c.setFont("Helvetica", 10)
            c.drawString(x, y, f"Spécialité : {spec}    Niveau : {level}    Groupe : {group}"); y -= 14
            dstr = now_dz().strftime("%d/%m/%Y %H:%M")
            c.drawString(x, y, f"Feuille de présence — Date/Heure : {dstr}"); y -= 8
            c.line(x, y, W - margin, y); y -= 16

            c.setFont("Helvetica-Bold", 10)
            c.drawString(x, y, "N°"); c.drawString(x+20, y, "Nom & Prénom")
            c.drawString(W - margin - 130, y, "Remarque")
            c.drawString(W - margin - 30, y, "Présent")
            y -= 12; c.line(x, y, W - margin, y); y -= 10
            c.setFont("Helvetica", 10)

            i = 1
            for name in stud["Nom affiché"].tolist():
                pres = st.session_state[key_state].get(name, False)
                remk = st.session_state[key_remark].get(name, "")
                if y < margin + 40:
                    c.showPage(); y = H - margin
                c.drawString(x, y, str(i)); c.drawString(x+20, y, name[:50])
                c.drawString(W - margin - 130, y, remk[:28])
                c.drawString(W - margin - 30, y, "✓" if pres else "✗")
                y -= 14; i += 1

            c.showPage(); c.save()
            buffer.seek(0); return buffer.getvalue()

        st.download_button(
            "📄 Exporter la présence en PDF",
            data=build_presence_pdf(),
            file_name=f"Presence_{spec}_{level}_{group}_{now_dz().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
# streamlit_app.py — PART 6/6

def main():
    mode, spec, level, group, search_name = filter_area()
    if not spec or not level or not group:
        st.info("Choisis Spécialité → Niveau → Groupe.")
        return
    if mode == "Étudiant":
        render_student_ui(spec, level, group, search_name)
    else:
        render_teacher_ui(spec, level, group, search_name)

if __name__ == "__main__":
    main()
