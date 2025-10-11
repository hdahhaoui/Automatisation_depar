"""Génère les fichiers maîtres utilisés par l'app Streamlit.

Ce script fusionne tous les classeurs d'emplois du temps situés dans
``data/raw/edt`` et, si disponibles, les listes d'étudiants placées dans
``data/raw/students``. Les fichiers générés sont enregistrés dans
``data/processed`` avec les noms attendus par l'application Streamlit :
``EDT_MASTER_S1.xlsx`` et ``ETUDIANTS_MASTER_S1.xlsx``.
"""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
RAW_EDT_DIR = DATA_DIR / "raw" / "edt"
RAW_ETU_DIR = DATA_DIR / "raw" / "students"
PROCESSED_DIR = DATA_DIR / "processed"

ORDER_JOUR = {
    "DIMANCHE": 0,
    "LUNDI": 1,
    "MARDI": 2,
    "MERCREDI": 3,
    "JEUDI": 4,
    "VENDREDI": 5,
    "SAMEDI": 6,
}


def _load_excels(directory: Path) -> Iterable[pd.DataFrame]:
    """Charge tous les classeurs Excel d'un dossier.

    Les fichiers sont parcourus par ordre alphabétique afin de garantir la
    reproductibilité du fichier fusionné.
    """

    if not directory.exists():
        return []

    frames = []
    for path in sorted(directory.glob("*.xls*")):
        df = pd.read_excel(path)
        df["__source"] = path.name
        frames.append(df)
    return frames


def build_edt_master() -> Optional[Path]:
    frames = _load_excels(RAW_EDT_DIR)
    if not frames:
        print("Aucun fichier EDT trouvé, saut du build EDT.")
        return None

    edt = pd.concat(frames, ignore_index=True)

    # Colonnes attendues par l'app.
    required_cols = [
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
    ]
    for col in required_cols:
        if col not in edt.columns:
            edt[col] = ""

    # Ordonnancement cohérent (jours de la semaine + heures).
    edt["__order"] = (
        edt["Jour"].astype(str).str.upper().map(ORDER_JOUR).fillna(99).astype(int)
    )
    edt = edt.sort_values([
        "Niveau",
        "Spécialité",
        "Groupe",
        "__order",
        "Heure début",
    ])
    edt = edt.drop(columns="__order")

    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    output_path = PROCESSED_DIR / "EDT_MASTER_S1.xlsx"
    edt.to_excel(output_path, index=False)
    print(f"Fichier EDT généré : {output_path}")
    return output_path


def build_students_master() -> Optional[Path]:
    frames = _load_excels(RAW_ETU_DIR)
    if not frames:
        # Crée un fichier vide avec les bonnes colonnes pour éviter les erreurs côté app.
        PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
        output_path = PROCESSED_DIR / "ETUDIANTS_MASTER_S1.xlsx"
        empty_cols = [
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
        pd.DataFrame(columns=empty_cols).to_excel(output_path, index=False)
        print(
            "Aucun fichier étudiants trouvé. Un classeur vide a été généré "
            f"à {output_path}"
        )
        return output_path

    etu = pd.concat(frames, ignore_index=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    output_path = PROCESSED_DIR / "ETUDIANTS_MASTER_S1.xlsx"
    etu.to_excel(output_path, index=False)
    print(f"Fichier étudiants généré : {output_path}")
    return output_path


def main() -> None:
    build_edt_master()
    build_students_master()


if __name__ == "__main__":
    main()
