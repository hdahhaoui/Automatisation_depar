# =========================
# app/streamlit_app.py
# =========================
# Thème sombre (facultatif) :
# crée .streamlit/config.toml :
# [theme]
# base="dark"
# primaryColor="#5eead4"
# backgroundColor="#0b0f17"
# secondaryBackgroundColor="#121826"
# textColor="#e5e7eb"

import re
import glob
from io import BytesIO
from datetime import datetime, timedelta, time as dtime
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
JOURS_FR = ["LUNDI","MARDI","MERCREDI","JEUDI","VENDREDI","SAMEDI","DIMANCHE"]

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

# ---------- Utilitaires ----------
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
    try:
        hh, mm = h.split("h")
        return int(hh or 0)*60 + int(mm or 0)
    except Exception:
        return None

def minutes_to_dt(d: datetime, minutes: int) -> datetime:
    return datetime.combine(d.date(), dtime.min) + timedelta(minutes=minutes)

def human_delta(dt: datetime, now: datetime):
    s = int((dt - now).total_seconds())
    d = s // 86400; s %= 86400
    h = s // 3600; s %= 3600
    m = s // 60
    out = []
    if d: out.append(f"{d}j")
    if h: out.append(f"{h}h")
    if m: out.append(f"{m}m")
    return " ".join(out) or "0m"

def next_session(now: datetime, edt_df: pd.DataFrame):
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
        dt = datetime.combine(day_date, dtime.min) + timedelta(minutes=m)
        if dt < now: dt += timedelta(days=7)
        rows.append((dt, r))
    if not rows: return None
    rows.sort(key=lambda x: x[0])
    return rows[0][0], rows[0][1]

# ---------- Normalisation ----------
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
    if s and not s.startswith("G") and s.isdigit():
        s = "G" + s
    return s

def classify_spec_level(spec_text: str, level_text: str):
    S = (spec_text or "").upper()
    L = (level_text or "").upper()
    if "RIB" in S: return "RIB", "M1" if "M1" in S+L else "M2" if "M2" in S+L else ""
    if "VOA" in S: return "VOA", "M1" if "M1" in S+L else "M2" if "M2" in S+L else ""
    if "STRUCT" in S: return "STRUCTURE", "M1" if "M1" in S+L else "M2" if "M2" in S+L else ""
    if "L2" in S+L: return "LICENCE","2"
    if "L3" in S+L: return "LICENCE","3"
    if "ING" in S+L:
        for n in ["1","2","3"]:
            if n in S+L: return "INGENIEUR", n
        return "INGENIEUR",""
    return "", ""

def infer_from_filename(path: str):
    name = Path(path).stem.upper().replace("-", "_")
    g = re.search(r"_G\s*?(\d+)", name)
    g = f"G{g.group(1)}" if g else None
    sem = re.search(r"_S\s*?(\d+)", name)
    sem = f"S{sem.group(1)}" if sem else "S1"
    if "RIB" in name: return "RIB", "M2" if "M2" in name else "M1", g, sem
    if "VOA" in name: return "VOA", "M2" if "M2" in name else "M1", g, sem
    if "STRUCT" in name: return "STRUCTURE", "M2" if "M2" in name else "M1", g, sem
    if "L2" in name: return "LICENCE","2",g,sem
    if "L3" in name: return "LICENCE","3",g,sem
    if "1ING" in name: return "INGENIEUR","1",g,sem
    if "2ING" in name: return "INGENIEUR","2",g,sem
    if "3ING" in name: return "INGENIEUR","3",g,sem
    return None,None,g,sem

def level_options_for(spec: str):
    if spec in ("RIB","VOA","STRUCTURE"): return ["M1","M2"]
    if spec=="LICENCE": return ["2","3"]
    if spec=="INGENIEUR": return ["1","2","3"]
    return []

def pretty_level_label(spec,niv):
    if spec=="LICENCE": return f"LICENCE {niv}"
    if spec=="INGENIEUR": return f"INGENIEUR {niv}"
    return niv

# ---------- Chargement ----------
@st.cache_data
def load_raw_s1():
    edt_list=[]; stu_list=[]
    for f in glob.glob(f"{RAW_EDT}/*_S1.*"):
        try:
            df=read_any(f)
            df=ensure_cols(df,EDT_COLS,["Durée (h)"])
            df["Semestre"]=df["Semestre"].apply(normalize_semestre)
            df["Groupe"]=df["Groupe"].apply(normalize_groupe)
            df["Spec2"],df["Niv2"]=zip(*df.apply(lambda r:classify_spec_level(r["Spécialité"],r["Niveau"]),axis=1))
            edt_list.append(df)
        except: pass
    edt=pd.concat(edt_list,ignore_index=True) if edt_list else pd.DataFrame(columns=EDT_COLS+["Spec2","Niv2"])
    for f in glob.glob(f"{RAW_STU}/*_S1.*"):
        try:
            df=read_any(f)
            df=ensure_cols(df,STU_COLS)
            df["Semestre"]=df["Semestre"].apply(normalize_semestre)
            df["Groupe"]=df["Groupe"].apply(normalize_groupe)
            df["Spec2"],df["Niv2"]=zip(*df.apply(lambda r:classify_spec_level(r["Spécialité"],r["Niveau"]),axis=1))
            stu_list.append(df)
        except: pass
    etu=pd.concat(stu_list,ignore_index=True) if stu_list else pd.DataFrame(columns=STU_COLS+["Spec2","Niv2"])
    return edt, etu

def subgroup(df,spec=None,niv=None,g=None):
    keep=df[df["Semestre"].astype(str).str.upper()==SEMESTRE]
    if spec: keep=keep[keep["Spec2"]==spec]
    if niv: keep=keep[keep["Niv2"]==niv]
    if g: keep=keep[keep["Groupe"].apply(normalize_groupe)==normalize_groupe(g)]
    return keep

# ---------------- Interface ----------------
st.title("🗓️ Portail Génie Civil — EDT & Listes (S1)")
edt, etu = load_raw_s1()
if edt.empty: st.stop()

with st.sidebar:
    st.subheader("🔎 Mode d’accès")
    role=st.radio("Je suis :",["Étudiant","Enseignant"],horizontal=True)
    st.markdown("---")
    spec_order=["RIB","VOA","STRUCTURE","LICENCE","INGENIEUR"]
    specs=[s for s in spec_order if s in edt["Spec2"].unique()]
    spec=st.selectbox("Spécialité",specs)
    raw=level_options_for(spec)
    labels=[pretty_level_label(spec,n) for n in raw]
    niv=raw[labels.index(st.selectbox("Niveau",labels))]
    grp=subgroup(edt,spec,niv)["Groupe"].dropna().unique()
    groupe=st.selectbox("Groupe",sorted(grp))
    q_nom=st.text_input("Nom/Prénom enseignant ou étudiant")
    print_mode=st.checkbox("🖨️ Mode impression")

bloc=subgroup(edt,spec,niv,groupe)
now=datetime.now()

# ---------------- Vue Enseignant ----------------
header_role("Enseignant", f"{spec} {niv} • Groupe {groupe}")
tab_plan,tab_next,tab_where,tab_presence=st.tabs(
    ["🗂️ Planning","⏭️ Prochaine séance","📍 Où trouver un enseignant ?","📝 Feuille de présence"]
)

# Planning
with tab_plan:
    view=bloc[["Jour","Heure début","Heure fin","Matière","Type","Salle","Groupe"]]
    st.download_button("⬇️ Exporter en Excel",df_to_xlsx_bytes(view),file_name=f"Planning_{spec}_{niv}_G{groupe}_S1.xlsx")
    st.dataframe(view,use_container_width=True,hide_index=True)

# Prochaine séance
with tab_next:
    nxt=next_session(now,bloc)
    if nxt:
        dt,r=nxt
        st.markdown(f"### {r['Matière']} — {r['Jour']} {r['Heure début']} à {r['Heure fin']}  \n📍 Salle {r['Salle']} — Groupe {r['Groupe']}  \n🕓 Dans {human_delta(dt,now)}")
    else:
        st.info("Aucune séance à venir.")

# Où trouver un enseignant
with tab_where:
    st.markdown("#### Où trouver un enseignant ?")
    only_today=st.checkbox("Aujourd’hui uniquement",value=False)
    base=edt.copy(); base=base[base["Semestre"]==SEMESTRE]
    base["Jour"]=base["Jour"].str.upper().str.strip()
    base["__start"]=base["Heure début"].map(time_to_minutes)
    base["__end"]=base["Heure fin"].map(time_to_minutes)
    today_idx=now.weekday(); today_name=JOURS_FR[today_idx]; now_min=now.hour*60+now.minute

    def next_occ(row):
        d_idx=ORDER_JOUR.get(row["Jour"])
        if d_idx is None:return None,False,False
        is_today=(d_idx==today_idx)
        if is_today and row["__start"]<=now_min<row["__end"]:return now,True,True
        delta=(d_idx-today_idx)%7;dt_day=(now+timedelta(days=delta)).date()
        dt=datetime.combine(dt_day,dtime.min)+timedelta(minutes=row["__start"])
        if is_today and row["__start"]<now_min:dt+=timedelta(days=7);is_today=False
        return dt,is_today,False

    rows=[]
    for ens,g in base.groupby("Enseignant"):
        if not ens:continue
        best=None
        for _,r in g.iterrows():
            dt,is_today,is_now=next_occ(r)
            if dt and (best is None or dt<best[0]):best=(dt,is_today,is_now,r)
        if not best:continue
        dt,is_today,is_now,r=best
        if is_now:
            stt=f"🟢 En cours jusqu’à {r['Heure fin']} (Salle {r['Salle']})";order=(0,dt)
        elif is_today:
            stt=f"🔵 Dans {human_delta(dt,now)}";order=(1,dt)
        else:
            stt=f"⚪ Le {r['Jour'].title()} à {r['Heure début']} (dans {human_delta(dt,now)})";order=(2,dt)
        rows.append({"Enseignant":ens,"Statut":stt,"Salle":r["Salle"],"Matière":r["Matière"],
                     "Jour":r["Jour"].title(),"Heure":r["Heure début"],"Groupe":r["Groupe"],
                     "Spécialité":r.get("Spec2",""),"Niveau":pretty_level_label(r.get("Spec2",""),r.get("Niv2","")),
                     "_o0":order[0],"_o1":order[1]})
    df=pd.DataFrame(rows)
    if only_today:df=df[df["Jour"].isin(["Aujourd’hui",today_name.title()])]
    if q_nom:df=df[df["Enseignant"].str.contains(q_nom,case=False,na=False)]
    if df.empty:st.info("Aucun enseignant trouvé.")
    else:st.dataframe(df.sort_values(["_o0","_o1","Enseignant"]).drop(columns=["_o0","_o1"]),use_container_width=True,hide_index=True)

# Feuille de présence
with tab_presence:
    etu_g=subgroup(etu,spec,niv,groupe)
    cols=[c for c in ["N°","Matricule","Nom","Prenom","Remarque"] if c in etu_g.columns]
    etu_g=etu_g[cols] if cols else etu_g
    if etu_g.empty:st.warning("Aucune liste trouvée.")
    else:
        key=f"pres_{spec}_{niv}_{groupe}"
        if key not in st.session_state:
            df=etu_g.copy();df["Présent"]=False;st.session_state[key]=df
        col1,col2,_=st.columns([1,1,2])
        if col1.button("Tout cocher"):st.session_state[key]["Présent"]=True
        if col2.button("Tout décocher"):st.session_state[key]["Présent"]=False
        edited=st.data_editor(st.session_state[key],num_rows="fixed",use_container_width=True)
        st.download_button("⬇️ Exporter la feuille en Excel",
                           df_to_xlsx_bytes(edited),
                           file_name=f"Présence_{spec}_{niv}_G{groupe}_S1.xlsx")
