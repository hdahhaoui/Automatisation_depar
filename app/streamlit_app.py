# ==============================
# Portail Génie Civil — EDT & Listes (S1)
# ==============================

from __future__ import annotations

import io
import os
import glob
import re
from datetime import datetime, timedelta, time
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta
import pytz

# PDF (feuilles de présence)
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import black

# ------------------------------
# CONFIG & CONSTANTES
# ------------------------------
st.set_page_config(
    page_title="Portail Génie Civil — EDT & Listes (S1)",
    page_icon="📅",
    layout="wide",
)

# Répertoires de données
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "raw")
EDT_DIR = os.path.join(DATA_DIR, "edt")
STUD_DIR = os.path.join(DATA_DIR, "students")

# Fuseau horaire Algérie
TZ = pytz.timezone("Africa/Algiers")

# Jours FR → ordre (utilisé pour boucler sur le prochain jour)
DAY_ORDER = ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]

# Jours FR pour affichage
DAY_LABELS = {
    "DIMANCHE": "DIMANCHE",
    "LUNDI": "LUNDI",
    "MARDI": "MARDI",
    "MERCREDI": "MERCREDI",
    "JEUDI": "JEUDI",
    "VENDREDI": "VENDREDI",
    "SAMEDI": "SAMEDI",
}

# Colonnes cibles EDT
EDT_COL_TARGETS = dict(
    jour="Jour",
    hstart="Heure début",
    hend="Heure fin",
    matiere="Matière",
    type="Type",
    ens="Enseignant",
    salle="Salle",
    freq="Fréquence",
)

# ------------------------------
# UTILITAIRES
# ------------------------------

def now_dz() -> datetime:
    """Datetime courant en TZ Algérie."""
    return datetime.now(tz=TZ)


def normalize_time_str(x: str) -> str:
    s = str(x).strip().replace(" ", "").lower()
    s = s.replace("h", ":")
    if ":" not in s:
        if s.isdigit():
            return f"{s}:00"
        return s
    # 8:30 → 08:30
    parts = s.split(":")
    if len(parts[0]) == 1:
        parts[0] = "0" + parts[0]
    if len(parts) == 1:
        parts.append("00")
    elif len(parts[1]) == 1:
        parts[1] = parts[1] + "0"
    return ":".join(parts[:2])


def parse_time(x: str) -> Optional[time]:
    try:
        s = normalize_time_str(x)
        hh, mm = s.split(":")
        return time(int(hh), int(mm))
    except Exception:
        return None


def jour_to_idx(j: str) -> int:
    j = (j or "").strip().upper()
    try:
        return DAY_ORDER.index(j)
    except ValueError:
        return -1


def idx_next_day(idx: int) -> int:
    return (idx + 1) % 7


def fmt_hhmm(t: time) -> str:
    return f"{t.hour:02d}h{t.minute:02d}"


def compact_timelapse(dt_from: datetime, dt_to: datetime) -> str:
    """Retourne 'Dans 2h35' ou 'En cours' ou 'Terminé'."""
    now = now_dz()
    if dt_from <= now <= dt_to:
        return "En cours"
    if now < dt_from:
        delta = dt_from - now
        h = delta.seconds // 3600
        m = (delta.seconds % 3600) // 60
        if h:
            return f"Dans {h}h{m:02d}"
        return f"Dans {m} min"
    return "Terminé"


def badge(label: str, color_bg: str = "#eaf2ff", color_fg: str = "#0842a0"):
    st.markdown(
        f"""
        <span style="display:inline-block;padding:6px 10px;border-radius:999px;
                     background:{color_bg};color:{color_fg};font-weight:600;font-size:0.85rem;
                     margin-right:8px;">{label}</span>
        """,
        unsafe_allow_html=True,
    )


# ------------------------------
# NORMALISATION (noms de fichiers EDT)
# ------------------------------

def norm_spec_from_filename(fname: str) -> Optional[str]:
    s = fname.upper()
    if "RIB" in s:
        return "RIB"
    if "VOA" in s:
        return "VOA"
    if "STR" in s or "STRUCTURE" in s:
        return "STRUCTURE"
    if "ING" in s or "INGENIEUR" in s:
        return "INGENIEUR"
    if "L2" in s or "L3" in s or "LICENCE" in s:
        return "LICENCE"
    return None


def norm_level_from_filename(fname: str) -> Optional[str]:
    s_raw = os.path.splitext(fname)[0].upper()
    compact = re.sub(r"[\s_\-]", "", s_raw)
    # Licence
    if "L2" in compact or "LICENCE2" in compact:
        return "LICENCE 2"
    if "L3" in compact or "LICENCE3" in compact:
        return "LICENCE 3"
    # Ingénieur
    if any(token in compact for token in ["1ING", "ING1", "INGENIEUR1"]):
        return "INGENIEUR 1"
    if any(token in compact for token in ["2ING", "ING2", "INGENIEUR2"]):
        return "INGENIEUR 2"
    if any(token in compact for token in ["3ING", "ING3", "INGENIEUR3"]):
        return "INGENIEUR 3"
    # Masters (M1/M2, MASTER1/MASTER2, M-1, etc.)
    token_str = re.sub(r"[_\-]", " ", s_raw)
    if (
        re.search(r"\bM\s*[-_]?\s*1\b", token_str)
        or "MASTER1" in compact
        or re.search(r"M1(?![0-9])", compact)
    ):
        return "M1"
    if (
        re.search(r"\bM\s*[-_]?\s*2\b", token_str)
        or "MASTER2" in compact
        or re.search(r"M2(?![0-9])", compact)
    ):
        return "M2"
    return None


def norm_group_from_filename(fname: str) -> Optional[str]:
    s = fname.upper()
    if "G11" in s:
        return "G11"
    if "G12" in s:
        return "G12"
    return None


# ------------------------------
# CHARGEMENT EDT
# ------------------------------

def map_edt_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Essaie de cartographier les colonnes de l'EDT vers les cibles standard."""
    mapping = {}
    for c in df.columns:
        cu = str(c).strip().lower()
        if "jour" in cu:
            mapping[c] = EDT_COL_TARGETS["jour"]
        elif ("début" in cu or "debut" in cu or "start" in cu) and "heure" in cu or cu == "debut":
            mapping[c] = EDT_COL_TARGETS["hstart"]
        elif ("fin" in cu or "end" in cu) and ("heure" in cu or cu == "fin"):
            mapping[c] = EDT_COL_TARGETS["hend"]
        elif "matière" in cu or "matiere" in cu:
            mapping[c] = EDT_COL_TARGETS["matiere"]
        elif cu.startswith("type"):
            mapping[c] = EDT_COL_TARGETS["type"]
        elif "enseignant" in cu or "prof" in cu:
            mapping[c] = EDT_COL_TARGETS["ens"]
        elif "salle" in cu:
            mapping[c] = EDT_COL_TARGETS["salle"]
        elif "fréquence" in cu or "frequence" in cu:
            mapping[c] = EDT_COL_TARGETS["freq"]

    df = df.rename(columns=mapping)
    # garde uniquement les colonnes cibles si présentes
    keep = [EDT_COL_TARGETS[k] for k in ["jour", "hstart", "hend", "matiere", "type", "ens", "salle", "freq"] if EDT_COL_TARGETS[k] in df.columns]
    df = df[keep].copy()

    # Normalise valeurs
    if "Jour" in df.columns:
        df["Jour"] = df["Jour"].astype(str).str.upper().str.strip()
    for col in ["Heure début", "Heure fin"]:
        if col in df.columns:
            df[col] = df[col].astype(str).map(lambda x: fmt_hhmm(parse_time(x)) if parse_time(x) else x)
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    return df


def default_levels_for_spec(spec: Optional[str]) -> List[str]:
    """Retourne la liste de niveaux par défaut attendus pour une spécialité."""
    spec = (spec or "").upper()
    mapping = {
        "INGENIEUR": ["INGENIEUR 1", "INGENIEUR 2", "INGENIEUR 3"],
        "LICENCE": ["LICENCE 2", "LICENCE 3"],
        "RIB": ["M1", "M2"],
        "VOA": ["M1", "M2"],
        "STRUCTURE": ["M1", "M2"],
    }
    return mapping.get(spec, ["M1", "M2"])


@st.cache_data(show_spinner=False)
def load_all_edt() -> Dict[Tuple[str, str, str], pd.DataFrame]:
    """Retourne {(spec, level, group): df_edt} pour toutes les feuilles EDT trouvées."""
    out: Dict[Tuple[str, str, str], pd.DataFrame] = {}
    for path in sorted(glob.glob(os.path.join(EDT_DIR, "*.xlsx"))):
        fname = os.path.basename(path)
        spec = norm_spec_from_filename(fname)
        level = norm_level_from_filename(fname)
        group = norm_group_from_filename(fname)

        try:
            raw = pd.read_excel(path)
        except Exception:
            continue

        df = map_edt_columns(raw)

        # Si jour non normalisé → essayer de déduire d'une 1ère colonne si liste vide
        if "Jour" in df.columns:
            df = df[df["Jour"].isin(DAY_ORDER) | df["Jour"].isin([d.capitalize() for d in DAY_ORDER])]

        if not spec or not level or not group:
            # Tentative heuristique supplémentaire
            text = " ".join(df.columns.astype(str)).upper() + " " + fname.upper()
            if not spec:
                spec = norm_spec_from_filename(text) or "INGENIEUR"
            if not level:
                detected = norm_level_from_filename(text)
                default_level = default_levels_for_spec(spec)[0]
                level = detected or default_level
            if not group:
                group = norm_group_from_filename(text) or "G11"

        out[(spec, level, group)] = df.reset_index(drop=True)
    return out


# ------------------------------
# CHARGEMENT LISTES ÉTUDIANTS (robuste)
# ------------------------------

@st.cache_data(show_spinner=False)
def load_all_students() -> Dict[Tuple[str,str,str], pd.DataFrame]:
    """
    Charge toutes les listes étudiants et normalise (Spécialité, Niveau, Groupe),
    d'abord via le nom de fichier, puis via le CONTENU (colonnes Spécialité/Niveau/Groupe)
    si le nom de fichier n'est pas assez parlant.
    """
    def norm_spec_from_value(val: str) -> Optional[str]:
        s = (val or "").strip().upper()
        if "RIB" in s: return "RIB"
        if "VOA" in s: return "VOA"
        if "STR" in s or "STRUCTURE" in s: return "STRUCTURE"
        if "LICENCE" in s or s in {"L2","L3"}: return "LICENCE"
        if "ING" in s or "INGENIEUR" in s or "GÉNIE CIVIL" in s or "GENIE CIVIL" in s: return "INGENIEUR"
        return None

    def norm_level_from_value(val: str) -> Optional[str]:
        s = (val or "").strip().upper().replace(" ", "")
        cleaned = re.sub(r"[\s_\-]", "", s)
        if cleaned in {"L2","LICENCE2"}: return "LICENCE 2"
        if cleaned in {"L3","LICENCE3"}: return "LICENCE 3"
        if cleaned in {"1ING","ING1","INGENIEUR1","1INGENIEUR"}: return "INGENIEUR 1"
        if cleaned in {"2ING","ING2","INGENIEUR2","2INGENIEUR"}: return "INGENIEUR 2"
        if cleaned in {"3ING","ING3","INGENIEUR3","3INGENIEUR"}: return "INGENIEUR 3"
        if cleaned in {"M1","MASTER1"}: return "M1"
        if cleaned in {"M2","MASTER2"}: return "M2"
        return None

    def norm_group_from_value(val: str) -> Optional[str]:
        s = (val or "").strip().upper()
        if "G11" in s: return "G11"
        if "G12" in s: return "G12"
        return None

    out: Dict[Tuple[str,str,str], pd.DataFrame] = {}

    for path in sorted(glob.glob(os.path.join(STUD_DIR, "*.xlsx"))):
        fname = os.path.basename(path)

        # 1) passe nom
        spec  = norm_spec_from_filename(fname)
        level = norm_level_from_filename(fname)
        group = norm_group_from_filename(fname)

        # 2) lecture
        try:
            df = pd.read_excel(path)
        except Exception:
            continue

        # 3) cartographie basique
        colmap = {}
        for c in df.columns:
            cu = str(c).strip().lower()
            if "nom" in cu and ("prénom" in cu or "prenom" in cu):
                colmap[c] = "NomComplet"
            elif cu.startswith("nom") and "prenom" not in cu and "prénom" not in cu:
                colmap[c] = "Nom"
            elif "prénom" in cu or "prenom" in cu:
                colmap[c] = "Prénom"
            elif "spécialité" in cu or "specialite" in cu:
                colmap[c] = "Col_Spécialité"
            elif "niveau" in cu:
                colmap[c] = "Col_Niveau"
            elif "groupe" in cu:
                colmap[c] = "Col_Groupe"
        df = df.rename(columns=colmap)

        # 4) passe contenu
        if spec is None and "Col_Spécialité" in df.columns:
            spec = norm_spec_from_value(df["Col_Spécialité"].dropna().astype(str).iloc[0])
        if level is None and "Col_Niveau" in df.columns:
            level = norm_level_from_value(df["Col_Niveau"].dropna().astype(str).iloc[0])
        if group is None and "Col_Groupe" in df.columns:
            group = norm_group_from_value(df["Col_Groupe"].dropna().astype(str).iloc[0])

        # 5) normalisation noms
        base = df.copy()
        base_cols = base.columns
        n = base.get("Nom") if "Nom" in base_cols else None
        p = base.get("Prénom") if "Prénom" in base_cols else None

        if n is not None:
            last_names = n.astype(str).fillna("").str.strip()
        else:
            last_names = pd.Series([""] * len(base), index=base.index, dtype="object")

        if p is not None:
            first_names = p.astype(str).fillna("").str.strip()
        else:
            first_names = pd.Series([""] * len(base), index=base.index, dtype="object")

        full_names = None
        if "NomComplet" in base_cols:
            full_names = base["NomComplet"].astype(str).fillna("").str.strip()

        if full_names is None or full_names.eq("").all():
            full_names = (last_names.str.strip() + " " + first_names.str.strip()).str.strip()

        # Si Nom/Prénom vides mais NomComplet renseigné → tenter un split simple
        if last_names.eq("").all() and first_names.eq("").all():
            split = full_names.str.split(r"\s+", n=1, expand=True)
            if split.shape[1] >= 1:
                last_names = split[0].fillna("").str.strip()
            if split.shape[1] >= 2:
                first_names = split[1].fillna("").str.strip()

        last_names = last_names.replace({"nan": ""}).fillna("").str.strip()
        first_names = first_names.replace({"nan": ""}).fillna("").str.strip()
        full_names = full_names.replace({"nan": ""}).fillna("").str.strip()

        mask_valid = (last_names != "") | (first_names != "") | (full_names != "")
        last_names = last_names[mask_valid]
        first_names = first_names[mask_valid]
        full_names = full_names[mask_valid]

        # 6) derniers recours
        if not spec or not level or not group:
            if not spec:
                text = " ".join(df.columns.astype(str)).upper()
                spec = "INGENIEUR" if "ING" in text else ("LICENCE" if "L" in text else None)
            if not level:
                text = df.get("Col_Niveau","").astype(str).str.upper().str.cat(sep=" ")
                if "1" in text and "ING" in text: level = "INGENIEUR 1"
                elif "2" in text and "ING" in text: level = "INGENIEUR 2"
                elif "3" in text and "ING" in text: level = "INGENIEUR 3"
                elif "L2" in text: level = "LICENCE 2"
                elif "L3" in text: level = "LICENCE 3"
                elif re.search(r"M\s*[-_]?\s*1|MASTER\s*1", text):
                    level = "M1"
                elif re.search(r"M\s*[-_]?\s*2|MASTER\s*2", text):
                    level = "M2"
            if not group:
                text = df.get("Col_Groupe","").astype(str).str.upper().str.cat(sep=" ")
                group = "G11" if "G11" in text else ("G12" if "G12" in text else None)

        if not spec or not level or not group:
            continue

        clean = pd.DataFrame({
            "Nom": last_names.tolist(),
            "Prénom": first_names.tolist(),
            "Nom complet": full_names.tolist(),
        }).reset_index(drop=True)
        out[(spec, level, group)] = clean

    return out


# ------------------------------
# UTIL: CONSTRUCTIONS D’OPTIONS / FILTRES
# ------------------------------

@st.cache_data(show_spinner=False)
def detect_filters(edt_dict: Dict[Tuple[str,str,str], pd.DataFrame]) -> Dict[str, Dict[str, List[str]]]:
    """
    Retourne une structure {spec: {levels:[], groups:[]}} à partir des EDT détectés.
    """
    res: Dict[str, Dict[str, List[str]]] = {}
    for (spec, level, group), df in edt_dict.items():
        res.setdefault(spec, dict(levels=set(), groups=set()))
        res[spec]["levels"].add(level)
        if group in {"G11", "G12"}:
            res[spec]["groups"].add(group)
    # listes triées
    out = {}
    for spec, d in res.items():
        out[spec] = dict(
            levels=sorted(d["levels"], key=lambda x: (x.startswith("INGENIEUR")*-1, x)),
            groups=sorted(d["groups"]),
        )
    return out


# ------------------------------
# RENDERS
# ------------------------------

def render_filters(filters_map: Dict[str, Dict[str, List[str]]]) -> Tuple[str, str, str]:
    """UI des filtres hiérarchiques : Spécialité → Niveau → Groupe."""
    specs = sorted(filters_map.keys(), key=lambda s: {"INGENIEUR": "0", "LICENCE":"1","RIB":"2","VOA":"3","STRUCTURE":"4"}.get(s, s))
    col = st.sidebar

    spec = col.selectbox("Spécialité", specs, index=0)

    levels = filters_map.get(spec, {}).get("levels", [])
    if not levels:
        levels = default_levels_for_spec(spec)
    level = col.selectbox("Niveau", levels, index=0)

    groups = filters_map.get(spec, {}).get("groups", [])
    if not groups:
        groups = ["G11", "G12"]
    group = col.selectbox("Groupe", groups, index=0)

    return spec, level, group


def edt_compact_view(df: pd.DataFrame) -> pd.DataFrame:
    """Affichage allégé pour mobile : Jour, Début, Fin, Matière, Enseignant, Salle."""
    keep = [c for c in ["Jour", "Heure début", "Heure fin", "Matière", "Enseignant", "Salle"] if c in df.columns]
    return df[keep].copy()


def export_df_to_xlsx_button(df: pd.DataFrame, filename: str, label: str):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    st.download_button(label, data=buf.getvalue(), file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def build_occurrence_datetime(day_name: str, hstart: str, hend: str) -> Tuple[datetime, datetime]:
    """Construit datetimes (début, fin) pour le prochain 'day_name' après maintenant."""
    now = now_dz()
    target_idx = jour_to_idx(day_name)
    if target_idx < 0:
        target_idx = DAY_ORDER.index(now.strftime("%A").upper())

    # combien de jours à ajouter pour atteindre ce jour
    today_idx = jour_to_idx(now.strftime("%A").upper())
    add = (target_idx - today_idx) % 7
    base_day = (now + timedelta(days=add)).date()

    t1 = parse_time(hstart) or time(8, 0)
    t2 = parse_time(hend) or time(9, 30)
    dt1 = TZ.localize(datetime.combine(base_day, t1))
    dt2 = TZ.localize(datetime.combine(base_day, t2))
    # Si le début est déjà passé aujourd'hui et qu'on est au même jour, garder quand même (pour "En cours")
    return dt1, dt2


def render_next_session(df: pd.DataFrame, title: str):
    st.subheader(title)

    # Jour par défaut = aujourd'hui (DZ)
    today_name = now_dz().strftime("%A").upper()
    today_name = {"SUNDAY":"DIMANCHE","MONDAY":"LUNDI","TUESDAY":"MARDI","WEDNESDAY":"MERCREDI","THURSDAY":"JEUDI","FRIDAY":"VENDREDI","SATURDAY":"SAMEDI"}.get(today_name, "DIMANCHE")
    all_days = [d for d in DAY_ORDER if d in (df["Jour"].unique().tolist() if "Jour" in df.columns else [])]
    if not all_days:
        st.info("Aucune séance trouvée pour ce filtre.")
        return

    day = st.selectbox("Jour", all_days, index=(all_days.index(today_name) if today_name in all_days else 0))

    ddf = df[df["Jour"] == day].copy()
    if ddf.empty:
        st.info("Aucune séance pour ce jour.")
        return

    # Trier par horaire
    if "Heure début" in ddf.columns:
        ddf = ddf.sort_values(by="Heure début", key=lambda s: s.map(lambda x: parse_time(x) or time(0,0)))

    now = now_dz()
    next_row = None
    current_row = None
    for _, r in ddf.iterrows():
        dt1, dt2 = build_occurrence_datetime(day, r.get("Heure début","08:00"), r.get("Heure fin","09:30"))
        if dt1 <= now <= dt2:
            current_row = r
            break
        if now < dt1 and next_row is None:
            next_row = r

    if current_row is not None:
        r = current_row
        dt1, dt2 = build_occurrence_datetime(day, r["Heure début"], r["Heure fin"])
        with st.container(border=True):
            st.markdown("**Séance en cours**")
            st.markdown(f"**{r.get('Matière','')} ({r.get('Type','')})**")
            st.caption(f"👤 {r.get('Enseignant','')}  •  🏫 Salle {r.get('Salle','')}  •  📅 {day}")
            st.caption(f"⏱️ {fmt_hhmm(parse_time(r.get('Heure début','08:00')))} – {fmt_hhmm(parse_time(r.get('Heure fin','09:30')))}")
    if next_row is not None:
        r = next_row
        dt1, dt2 = build_occurrence_datetime(day, r["Heure début"], r["Heure fin"])
        with st.container(border=True):
            st.markdown("**Après**")
            st.markdown(f"**{r.get('Matière','')} ({r.get('Type','')})**")
            st.caption(f"👤 {r.get('Enseignant','')}  •  🏫 Salle {r.get('Salle','')}  •  📅 {day}")
            st.caption(f"⏳ {compact_timelapse(dt1, dt2)} (de {fmt_hhmm(parse_time(r.get('Heure début','08:00')))} à {fmt_hhmm(parse_time(r.get('Heure fin','09:30')))})")
    if current_row is None and next_row is None:
        st.info("Tous les créneaux de ce jour sont passés.")

def render_where_is_teacher(edt_dict: Dict[Tuple[str,str,str], pd.DataFrame], spec: str, level: str, group: str):
    st.subheader("Où trouver un enseignant ?")

    # Jour courant
    today_name = now_dz().strftime("%A").upper()
    today_name = {"SUNDAY":"DIMANCHE","MONDAY":"LUNDI","TUESDAY":"MARDI","WEDNESDAY":"MERCREDI","THURSDAY":"JEUDI","FRIDAY":"VENDREDI","SATURDAY":"SAMEDI"}.get(today_name, "DIMANCHE")

    # Fusion EDT de la spécialité/niveau (tous groupes confondus pour la recherche)
    frames = []
    for (sp, lv, gp), df in edt_dict.items():
        if sp == spec and lv == level:
            frames.append(df)
    if not frames:
        st.info("Aucun EDT pour cette combinaison.")
        return
    big = pd.concat(frames, ignore_index=True)
    if "Jour" not in big.columns:
        st.info("Format EDT incomplet (colonne 'Jour' manquante).")
        return

    day = st.selectbox("Jour", sorted(big["Jour"].unique().tolist(), key=lambda d: DAY_ORDER.index(d) if d in DAY_ORDER else 999),
                       index= (sorted(big["Jour"].unique().tolist()).index(today_name) if today_name in big["Jour"].unique() else 0))

    search = st.text_input("Nom d'enseignant (laisser vide pour tous)", "")
    ddf = big[big["Jour"] == day].copy()

    if search.strip():
        mask = ddf["Enseignant"].str.contains(search.strip(), case=False, na=False)
        ddf = ddf[mask]

    if ddf.empty:
        st.info("Aucun cours planifié pour ce jour / filtre.")
        return

    # Calcul prochaine séance par enseignant
    now = now_dz()
    out_rows = []
    for ens, part in ddf.groupby("Enseignant"):
        part = part.sort_values(by="Heure début", key=lambda s: s.map(lambda x: parse_time(x) or time(0,0)))
        next_slot = None
        for _, r in part.iterrows():
            dt1, dt2 = build_occurrence_datetime(day, r["Heure début"], r["Heure fin"])
            if now <= dt2:
                next_slot = r
                break
        if next_slot is None:
            continue
        out_rows.append(dict(
            Enseignant=ens,
            Prochaine=f"{next_slot.get('Matière','')} ({next_slot.get('Type','')})",
            Salle=next_slot.get("Salle",""),
            Horaire=f"{next_slot.get('Heure début','')}–{next_slot.get('Heure fin','')}",
        ))

    if not out_rows:
        st.info("Aucun cours restant pour ce jour.")
        return

    tab = pd.DataFrame(out_rows)
    st.dataframe(tab, hide_index=True, use_container_width=True)


# ------------------------------
# FEUILLE DE PRÉSENCE (enseignant)
# ------------------------------

def pdf_presence(spec: str, level: str, group: str, df_presence: pd.DataFrame) -> bytes:
    """Construit un PDF (bytes) de la feuille de présence avec remarques."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # En-tête
    c.setFont("Helvetica-Bold", 13)
    c.drawString(2 * cm, h - 2 * cm, "Université de Tlemcen")
    c.setFont("Helvetica", 11)
    c.drawString(2 * cm, h - 2.7 * cm, "Faculté de Technologie")
    c.drawString(2 * cm, h - 3.4 * cm, "Département de Génie Civil")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(2 * cm, h - 4.4 * cm, f"Spécialité : {spec}   •   Niveau : {level}   •   Groupe : {group}")
    dt = now_dz().strftime("%d/%m/%Y %H:%M")
    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, h - 5.1 * cm, f"Date / Heure : {dt}")

    # Tableau : Nom | Prénom | Présent | Remarque
    y = h - 6 * cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2 * cm, y, "Nom")
    c.drawString(7 * cm, y, "Prénom")
    c.drawString(11 * cm, y, "Présent")
    c.drawString(14 * cm, y, "Remarque")
    y -= 0.6 * cm
    c.setFont("Helvetica", 10)

    for _, r in df_presence.iterrows():
        nom = str(r.get("Nom", "") or "")
        prenom = str(r.get("Prénom", "") or "")
        if not nom and not prenom:
            nom = str(r.get("Nom complet", "") or "")
        present = "Oui" if r["Présent"] else "Non"
        rem = str(r.get("Remarque","") or "")
        if y < 2 * cm:
            c.showPage()
            y = h - 2 * cm
        c.drawString(2 * cm, y, nom[:30])
        c.drawString(7 * cm, y, prenom[:30])
        c.drawString(11 * cm, y, present)
        c.drawString(14 * cm, y, rem[:35])
        y -= 0.5 * cm

    c.showPage()
    c.save()
    return buf.getvalue()


def render_presence(spec: str, level: str, group: str, students_map: Dict[Tuple[str,str,str], pd.DataFrame]):
    st.subheader("Feuille de présence (enseignant)")

    stud = students_map.get((spec, level, group))
    if stud is None or stud.empty:
        st.warning("Aucune liste d'étudiants détectée pour ce groupe.")
        return

    stud = stud.copy().reset_index(drop=True)
    for col in ["Nom", "Prénom", "Nom complet"]:
        if col in stud.columns:
            stud[col] = stud[col].astype(str).replace({"nan": ""}).fillna("").str.strip()

    if "Nom" not in stud.columns:
        stud["Nom"] = ""
    if "Prénom" not in stud.columns:
        stud["Prénom"] = ""
    if "Nom complet" not in stud.columns:
        stud["Nom complet"] = (stud["Nom"].str.strip() + " " + stud["Prénom"].str.strip()).str.strip()

    # assure une valeur d'affichage cohérente
    stud["Nom complet"] = stud["Nom complet"].replace({"nan": ""}).fillna("").str.strip()
    empty_full = stud["Nom complet"] == ""
    stud.loc[empty_full, "Nom complet"] = (stud.loc[empty_full, "Nom"].str.strip() + " " + stud.loc[empty_full, "Prénom"].str.strip()).str.strip()

    stud["__id"] = stud.index.map(lambda i: f"{spec}_{level}_{group}_{i}")

    # Mode mobile compact
    mobile = st.toggle("📱 Mode mobile (affichage compact)", value=True, help="Nom + case Présent (pas de défilement horizontal).")

    # Recherche rapide
    q = st.text_input("🔎 Recherche rapide (Nom/Prénom) :", "")
    show = stud.copy()
    if q.strip():
        needle = q.strip()
        mask = (
            show["Nom"].str.contains(needle, case=False, na=False)
            | show["Prénom"].str.contains(needle, case=False, na=False)
            | show["Nom complet"].str.contains(needle, case=False, na=False)
        )
        show = show[mask]

    # État session : cases + remarques
    key_state = f"presence_{spec}_{level}_{group}"
    if key_state not in st.session_state:
        st.session_state[key_state] = {row["__id"]: False for _, row in stud.iterrows()}
    key_rem = f"remark_{spec}_{level}_{group}"
    if key_rem not in st.session_state:
        st.session_state[key_rem] = {row["__id"]: "" for _, row in stud.iterrows()}

    # Flags pour rerun propre
    if "__needs_rerun__" not in st.session_state:
        st.session_state["__needs_rerun__"] = False

    def set_all(val: bool):
        for ident in show["__id"].tolist():
            st.session_state[key_state][ident] = val
            st.session_state[(key_state, ident)] = val
        st.session_state["__needs_rerun__"] = True

    c1, c2 = st.columns(2)
    with c1:
        st.button("✅ Tout cocher", use_container_width=True, on_click=set_all, args=(True,))
    with c2:
        st.button("❌ Tout décocher", use_container_width=True, on_click=set_all, args=(False,))

    if st.session_state.get("__needs_rerun__", False):
        st.session_state["__needs_rerun__"] = False
        st.rerun()

    # Affichage liste
    table_rows = []
    if mobile:
        # Liste compacte (Nom + ✅)
        st.caption("En mode mobile : Nom + case Présent. Les remarques restent disponibles plus bas.")
        for _, row in show.iterrows():
            ident = row["__id"]
            nom = row["Nom"].strip()
            prenom = row["Prénom"].strip()
            display_name = (nom + " " + prenom).strip() or row["Nom complet"]
            colA, colB = st.columns([4, 1])
            with colA:
                st.write(display_name)
            with colB:
                chk = st.checkbox(
                    "Présent",
                    key=(key_state, ident),
                    value=st.session_state[key_state].get(ident, False),
                )
                st.session_state[key_state][ident] = st.session_state.get((key_state, ident), chk)
        with st.expander("✍️ Remarques (facultatif)", expanded=False):
            for _, row in show.iterrows():
                ident = row["__id"]
                display_name = (row["Nom"].strip() + " " + row["Prénom"].strip()).strip() or row["Nom complet"]
                txt = st.text_input(
                    display_name or "Étudiant",
                    key=(key_rem, ident),
                    value=st.session_state[key_rem].get(ident, ""),
                )
                st.session_state[key_rem][ident] = st.session_state.get((key_rem, ident), txt)
    else:
        # Tableau avec présence + remarque
        grid = []
        for _, row in show.iterrows():
            ident = row["__id"]
            nom = row["Nom"].strip()
            prenom = row["Prénom"].strip()
            display_name = (nom + " " + prenom).strip() or row["Nom complet"]
            present = st.session_state[key_state].get(ident, False)
            remark = st.session_state[key_rem].get(ident, "")
            c1, c2, c3 = st.columns([4, 1, 3])
            with c1:
                st.write(display_name)
            with c2:
                chk = st.checkbox("Présent", key=(key_state, ident), value=present)
                st.session_state[key_state][ident] = st.session_state.get((key_state, ident), chk)
            with c3:
                txt = st.text_input(
                    "Remarque",
                    key=(key_rem, ident),
                    value=remark,
                    label_visibility="collapsed",
                )
                st.session_state[key_rem][ident] = st.session_state.get((key_rem, ident), txt)

    # Construire DataFrame présence
    pres_rows = []
    for _, row in show.iterrows():
        ident = row["__id"]
        pres_rows.append({
            "Nom": row["Nom"].strip(),
            "Prénom": row["Prénom"].strip(),
            "Nom complet": row["Nom complet"].strip(),
            "Présent": bool(st.session_state[key_state].get(ident, False)),
            "Remarque": st.session_state[key_rem].get(ident, ""),
        })
    df_presence = pd.DataFrame(pres_rows)

    # Export PDF
    pdf = pdf_presence(spec, level, group, df_presence)
    st.download_button("🧾 Exporter la présence en PDF", data=pdf, file_name=f"Presence_{spec}_{level}_{group}.pdf", mime="application/pdf")


# ------------------------------
# PAGES ÉTUDIANT / ENSEIGNANT
# ------------------------------

def render_student_ui(edt_df: pd.DataFrame, spec: str, level: str, group: str):
    badge("🎓 Étudiant", "#E6F6EE", "#0A7C5B")
    badge(f"{spec} {level} • Groupe {group}", "#EEF2FF", "#1F4B99")

    tab1, tab2 = st.tabs(["📘 Mon EDT", "⏱️ Prochaine séance"])
    with tab1:
        if edt_df is None or edt_df.empty:
            st.info("EDT vide pour ce filtre.")
        else:
            df_show = edt_compact_view(edt_df)
            export_df_to_xlsx_button(df_show, f"EDT_{spec}_{level}_{group}.xlsx", "📥 Exporter l’EDT en Excel")
            st.dataframe(df_show, hide_index=True, use_container_width=True)
    with tab2:
        render_next_session(edt_df, "Prochaine séance (Étudiant)")


def render_teacher_ui(edt_df: pd.DataFrame, spec: str, level: str, group: str, edt_dict: Dict, students_map: Dict):
    badge("👨‍🏫 Enseignant", "#FFF4E5", "#8A4F00")
    badge(f"{spec} {level} • Groupe {group}", "#EEF2FF", "#1F4B99")

    tab1, tab2, tab3 = st.tabs(["📘 Planning", "⏱️ Prochaine séance", "🧾 Feuille de présence"])
    with tab1:
        if edt_df is None or edt_df.empty:
            st.info("EDT vide pour ce filtre.")
        else:
            df_show = edt_compact_view(edt_df)
            export_df_to_xlsx_button(df_show, f"Planning_{spec}_{level}_{group}.xlsx", "📥 Exporter l’EDT en Excel")
            st.dataframe(df_show, hide_index=True, use_container_width=True)
    with tab2:
        render_next_session(edt_df, "Prochaine séance (Enseignant)")
    with tab3:
        render_presence(spec, level, group, students_map)


# ------------------------------
# MAIN
# ------------------------------

def main():
    st.title("📅 Portail Génie Civil — EDT & Listes (S1)")

    # Chargement
    edt_map = load_all_edt()
    stud_map = load_all_students()
    filters_map = detect_filters(edt_map)

    # Barre latérale : mode d’accès
    st.sidebar.header("🔎 Mode d’accès")
    mode = st.sidebar.radio("Je suis :", ["Étudiant", "Enseignant"], index=0)

    # Filtres hiérarchiques
    spec, level, group = render_filters(filters_map)

    # EDT du groupe
    edt_df = edt_map.get((spec, level, group), pd.DataFrame())

    # Navigation additionnelle : “Où trouver un enseignant ?” (spécialité/niveau)
    st.sidebar.markdown("---")
    if st.sidebar.button("📍 Où trouver un enseignant ? (vue du jour)"):
        st.session_state["show_where"] = True
    if st.session_state.get("show_where", False):
        with st.expander("📍 Où trouver un enseignant ? — Vue hebdo du jour (spécialité/niveau)", expanded=True):
            render_where_is_teacher(edt_map, spec, level, group=None)
        st.session_state["show_where"] = False

    # Corps de page
    if mode == "Étudiant":
        render_student_ui(edt_df, spec, level, group)
    else:
        render_teacher_ui(edt_df, spec, level, group, edt_map, stud_map)

    st.caption("S1 • Spécialité → Niveau → Groupe • Groupes normalisés (G11/G12) • Harmonisation des listes étudiants • Exports en Excel (EDT) & PDF (présence) • Fuseau horaire : Afrique/Alger.")


if __name__ == "__main__":
    main()
