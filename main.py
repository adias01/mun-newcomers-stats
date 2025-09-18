import csv
import os
import sys

from datetime import date
from pathlib import Path
from tabulate import tabulate

if getattr(sys, 'frozen', False): # Running as executable
    BASE_DIR = Path(sys.executable).parent
else:  # Running as script
    BASE_DIR = Path(__file__).parent

cwd_text_files = BASE_DIR / "text_files"
cwd_csv_files = BASE_DIR / "csv_files"

cwd_text_files.mkdir(exist_ok=True)
cwd_csv_files.mkdir(exist_ok=True)


def validate_txtfile(file_name, file_ext):
    if file_ext.lower() != ".txt":
        sys.exit(f"ERROR: {file_name}{file_ext} is not a text file")

    try:
        with open(cwd_text_files / f"{file_name}{file_ext}", encoding="utf-8") as f:
            f.readable()
    except OSError:
        sys.exit(f"ERROR: {file_name}{file_ext} could not be read")


def convert_to_csv(file_name, file_ext):
    txtfile = cwd_text_files / f"{file_name}{file_ext}"
    file_path = cwd_csv_files / f"{Path(file_name).stem}.csv"

    with open(txtfile, encoding="utf-8") as txt, open(
        file_path, "w", encoding="utf-8", newline=""
    ) as csvf:
        reader = txt.readlines()
        if reader:
            reader.pop(0) # Removes header

        fieldnames = [
            "f.aar", "navn", "adresse1", "adresse2",
            "postnr", "poststed", "flyttet", "tidl.knr", "tidl.k"
        ]
        writer = csv.DictWriter(csvf, fieldnames=fieldnames)
        writer.writeheader()

        for row in reader:
            newcomer = row.strip().split(";")
            # Replaces empty fields
            newcomer = [item if item else "N/A" for item in newcomer]

            writer.writerow({
                "f.aar": newcomer[0],
                "navn": newcomer[1],
                "adresse1": newcomer[2],
                "adresse2": newcomer[3],
                "postnr": newcomer[4],
                "poststed": newcomer[5],
                "flyttet": newcomer[6],
                "tidl.knr": newcomer[7],
                "tidl.k": newcomer[8] if len(newcomer) > 8 else "N/A"
            })


def age_stats(file_name):
    file_path = cwd_csv_files / f"{file_name}.csv"
    with open(file_path, encoding="utf-8") as file:
        today = date.today()
        reader = csv.DictReader(file)

        counts = {"0-5": 0, "6-19": 0, "20-44": 0, "45-66": 0, "67+": 0}
        for row in reader:
            age = today.year - int(row["f.aar"])
            if 0 <= age <= 5:
                counts["0-5"] += 1
            elif 6 <= age <= 19:
                counts["6-19"] += 1
            elif 20 <= age <= 44:
                counts["20-44"] += 1
            elif 45 <= age <= 66:
                counts["45-66"] += 1
            elif age >= 67:
                counts["67+"] += 1

        total = sum(counts.values())

    return total, counts


def pre_mun_stats(file_name):
    file_path = cwd_csv_files / f"{file_name}.csv"
    with open(file_path, encoding="utf-8") as file:
        reader = csv.DictReader(file)

        counts = {"stange": 0, "ringsaker": 0, "løten": 0, "elverum": 0, "oslo": 0}
        for row in reader:
            mun_no = row["tidl.knr"]
            if mun_no == "3413":
                counts["stange"] += 1
            elif mun_no == "3411":
                counts["ringsaker"] += 1
            elif mun_no == "3412":
                counts["løten"] += 1
            elif mun_no == "3420":
                counts["elverum"] += 1
            elif mun_no == "0301":
                counts["oslo"] += 1

    return counts


def month_stats(file_name):
    months = {
        1: "januar", 2: "februar", 3: "mars", 4: "april", 5: "mai", 6: "juni",
        7: "juli", 8: "august", 9: "september", 10: "oktober",
        11: "november", 12: "desember"
    }
    month_no = int(file_name[4:6]) - 1
    if month_no == 0:
        month_no = 12
    return months.get(month_no, "ukjent")


def main():
    stat_file = BASE_DIR / "statistikk_innflyttere.csv"
    fieldnames = [
        "måned", "antall innflyttere", "0-5", "6-19", "20-44", "45-66", "67+",
        "stange", "ringsaker", "løten", "elverum", "oslo"
    ]

    with open(stat_file, "w", encoding="utf-8", newline="") as csvf:
        writer = csv.DictWriter(csvf, fieldnames=fieldnames)
        writer.writeheader()

        for file in cwd_text_files.glob("*.txt"):
            file_name, file_ext = os.path.splitext(file.name)
            validate_txtfile(file_name, file_ext)
            convert_to_csv(file_name, file_ext)

            month = month_stats(file_name)
            total, ages = age_stats(file_name)
            mun_counts = pre_mun_stats(file_name)

            writer.writerow({
                "måned": month,
                "antall innflyttere": total,
                **ages,
                **mun_counts
            })


    with open(stat_file, encoding="utf-8") as csvf:
        reader = csv.DictReader(csvf)
        stat_dict = [row for row in reader]
        table = tabulate(stat_dict, headers="keys", tablefmt="grid")

    txt_output = BASE_DIR / "statistikk_innflyttere.txt"
    with open(txt_output, "w", encoding="utf-8") as f:
        f.write(table)

    print("\n", table, "\n")



if __name__ == "__main__":
    main()

