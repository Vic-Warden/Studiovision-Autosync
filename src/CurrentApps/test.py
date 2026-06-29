import os
import re
import time
from datetime import datetime
from pathlib import Path

# Racine à analyser
DEST_PHOTOS = Path(r"M:\PHOTOS")

# Dossier de sortie du rapport (Bureau de l'utilisateur courant)
OUTPUT_DIR = Path(os.path.join(os.path.expanduser("~"), "Desktop"))

# Reconnaît les dossiers parents : "-1.000" à "-9.000" et "00.000" à "99.000"
PARENT_FOLDER_RE = re.compile(r"^-?\d{1,2}\.000$")

# Si le scan d'un seul dossier parent prend plus que ce délai (en secondes),
# on affiche un avertissement explicite pour repérer un dossier réseau lent.
SLOW_PARENT_WARNING_SECONDS = 15


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


def log(msg: str) -> None:
    """Affiche un message immédiatement (sans tampon) avec l'heure."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}", flush=True)


def find_duplicate_patient_folders(root: Path) -> dict[str, list[Path]]:
    """
    Parcourt tous les dossiers parents (XX.000 / -X.000) sous root, regroupe
    leurs sous-dossiers par code patient, et retourne uniquement les codes
    ayant plusieurs dossiers DIFFÉRENTS (doublons).
    Affiche la progression en temps réel.
    """
    by_code: dict[str, list[Path]] = {}

    log(f"Lecture du contenu de {root} ...")
    t0 = time.monotonic()
    try:
        parent_entries = sorted(root.iterdir())
    except Exception as exc:
        log(f"ERREUR: impossible de lister {root}: {exc}")
        return {}
    log(f"  -> {len(parent_entries)} entrée(s) trouvée(s) en {time.monotonic() - t0:.1f}s")

    parent_folders = [p for p in parent_entries if p.is_dir() and PARENT_FOLDER_RE.match(p.name)]
    log(f"  -> {len(parent_folders)} dossier(s) parent(s) reconnu(s) (XX.000 / -X.000)")

    total_patients = 0
    scan_start = time.monotonic()

    for i, parent in enumerate(parent_folders, start=1):
        t_parent_start = time.monotonic()
        log(f"[{i}/{len(parent_folders)}] Scan de {parent.name} ...")

        try:
            patient_entries = list(parent.iterdir())
        except Exception as exc:
            log(f"  ATTENTION: impossible de lister {parent}: {exc}")
            continue

        count_in_parent = 0
        for patient_folder in patient_entries:
            if not patient_folder.is_dir():
                continue
            code = extract_code(patient_folder.name)
            if code is None:
                continue
            by_code.setdefault(code, []).append(patient_folder)
            count_in_parent += 1

        total_patients += count_in_parent
        elapsed_parent = time.monotonic() - t_parent_start
        log(f"  -> {count_in_parent} dossier(s) patient(s) en {elapsed_parent:.1f}s "
            f"(total cumulé : {total_patients})")

        if elapsed_parent > SLOW_PARENT_WARNING_SECONDS:
            log(f"  !!! Ce dossier a été anormalement lent à lire ({elapsed_parent:.1f}s) "
                f"— possible lenteur réseau sur {parent}")

    total_elapsed = time.monotonic() - scan_start
    log(f"Scan terminé : {total_patients} dossier(s) patient(s) au total en {total_elapsed:.1f}s")

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
    log(f"Analyse de {DEST_PHOTOS} ...")

    if not DEST_PHOTOS.is_dir():
        log(f"ERREUR: dossier introuvable ou inaccessible : {DEST_PHOTOS}")
        log("Vérifiez que le lecteur réseau M: est bien connecté "
            "(ouvrez l'explorateur de fichiers et essayez d'accéder à M:\\PHOTOS manuellement).")
        input("Appuyez sur Entrée pour fermer...")
        return

    duplicates = find_duplicate_patient_folders(DEST_PHOTOS)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    output_path = OUTPUT_DIR / f"doublons_patients_{timestamp}.txt"

    try:
        write_report(duplicates, output_path)
    except Exception as exc:
        log(f"ERREUR: impossible d'écrire le rapport: {exc}")
        input("Appuyez sur Entrée pour fermer...")
        return

    log(f"Terminé. {len(duplicates)} code(s) patient en doublon trouvé(s).")
    log(f"Rapport écrit dans : {output_path}")

    try:
        os.startfile(str(output_path))
    except Exception:
        pass

    input("Appuyez sur Entrée pour fermer...")


if __name__ == "__main__":
    main()