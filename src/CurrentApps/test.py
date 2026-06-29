import os
import re
from datetime import datetime
from pathlib import Path

# Racine à analyser
DEST_PHOTOS = Path(r"M:\PHOTOS")

# Dossier de sortie du rapport (Bureau de l'utilisateur courant)
OUTPUT_DIR = Path(os.path.join(os.path.expanduser("~"), "Desktop"))

# Reconnaît les dossiers parents : "-1.000" à "-9.000" et "00.000" à "99.000"
PARENT_FOLDER_RE = re.compile(r"^-?\d{1,2}\.000$")

# Extrait le code patient en tête du nom de dossier, en ignorant un éventuel
# "-" parasite. Le code est la suite de chiffres consécutifs au début.
CODE_PREFIX_RE = re.compile(r"^-?(\d+)")


def extract_code(folder_name: str) -> str | None:
    """
    Retourne le code patient (suite de chiffres en tête, "-" parasite ignoré)
    d'un nom de dossier, ou None si aucun chiffre n'est trouvé en tête.
    """
    name_without_dash = folder_name.lstrip("-")
    match = re.match(r"^(\d+)", name_without_dash)
    if not match:
        return None
    return match.group(1)


def find_duplicate_patient_folders(root: Path) -> dict[str, list[Path]]:
    """
    Parcourt tous les dossiers parents (XX.000 / -X.000) sous root, regroupe
    leurs sous-dossiers par code patient, et retourne uniquement les codes
    ayant plusieurs dossiers DIFFÉRENTS (doublons).
    """
    by_code: dict[str, list[Path]] = {}

    try:
        parent_entries = sorted(root.iterdir())
    except Exception as exc:
        print(f"ERREUR: impossible de lister {root}: {exc}")
        return {}

    for parent in parent_entries:
        if not parent.is_dir():
            continue
        if not PARENT_FOLDER_RE.match(parent.name):
            continue

        try:
            patient_entries = parent.iterdir()
        except Exception as exc:
            print(f"  ATTENTION: impossible de lister {parent}: {exc}")
            continue

        for patient_folder in patient_entries:
            if not patient_folder.is_dir():
                continue
            code = extract_code(patient_folder.name)
            if code is None:
                continue
            by_code.setdefault(code, []).append(patient_folder)

    # On ne garde que les codes ayant 2+ dossiers DONT LES NOMS DIFFÈRENT
    duplicates: dict[str, list[Path]] = {}
    for code, folders in by_code.items():
        distinct_names = {f.name for f in folders}
        if len(distinct_names) > 1:
            duplicates[code] = sorted(folders, key=lambda p: p.name)

    return duplicates


def write_report(duplicates: dict[str, list[Path]], output_path: Path) -> None:
    lines: list[str] = []
    lines.append("=" * 70)
    lines.append("RAPPORT DE DOSSIERS PATIENTS EN DOUBLON")
    lines.append(f"Généré le : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Racine analysée : {DEST_PHOTOS}")
    lines.append(f"Nombre de codes patients en doublon : {len(duplicates)}")
    lines.append("=" * 70)
    lines.append("")

    if not duplicates:
        lines.append("Aucun doublon trouvé.")
    else:
        for code in sorted(duplicates.keys(), key=lambda c: (len(c), c)):
            folders = duplicates[code]
            lines.append(f"Code patient : {code}  ({len(folders)} dossiers)")
            for folder in folders:
                try:
                    nb_files = sum(1 for f in folder.iterdir() if f.is_file())
                except Exception:
                    nb_files = "?"
                lines.append(f"    - {folder}  [{nb_files} fichier(s)]")
            lines.append("")

    output_path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    print(f"Analyse de {DEST_PHOTOS} ...")

    if not DEST_PHOTOS.is_dir():
        print(f"ERREUR: dossier introuvable ou inaccessible : {DEST_PHOTOS}")
        input("Appuyez sur Entrée pour fermer...")
        return

    duplicates = find_duplicate_patient_folders(DEST_PHOTOS)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    output_path = OUTPUT_DIR / f"doublons_patients_{timestamp}.txt"

    try:
        write_report(duplicates, output_path)
    except Exception as exc:
        print(f"ERREUR: impossible d'écrire le rapport: {exc}")
        input("Appuyez sur Entrée pour fermer...")
        return

    print(f"Terminé. {len(duplicates)} code(s) patient en doublon trouvé(s).")
    print(f"Rapport écrit dans : {output_path}")

    try:
        os.startfile(str(output_path))
    except Exception:
        pass

    input("Appuyez sur Entrée pour fermer...")


if __name__ == "__main__":
    main()