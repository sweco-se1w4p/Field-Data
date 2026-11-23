from pathlib import Path
import re
from openpyxl import load_workbook

# Mappen där denna .py-fil ligger
mapp = Path(__file__).resolve().parent

# Skapar undermappen "TXT filer" om den inte redan finns
txt_mapp = mapp / "TXT filer"
txt_mapp.mkdir(exist_ok=True)

# Filnamn utan filändelse
filnamn = "BULLERFORSEN_20261102 - kopia"

# Letar efter .xlsx eller .xlsm
excel_filer = list(mapp.glob(f"{filnamn}.xlsx")) + list(mapp.glob(f"{filnamn}.xlsm"))

if not excel_filer:
    raise FileNotFoundError(
        f"Hittar inte '{filnamn}.xlsx' eller '{filnamn}.xlsm' i:\n{mapp}"
    )

excel_fil = excel_filer[0]

# data_only=True = läs det beräknade värdet, inte formeln
wb = load_workbook(
    excel_fil,
    data_only=True,
    read_only=True
)

if "README" not in wb.sheetnames:
    raise ValueError("Fliken 'README' hittades inte i arbetsboken.")

# Bara flikarna som ligger efter README
readme_index = wb.sheetnames.index("README")
flikar_att_exportera = wb.sheetnames[readme_index + 1:]

if not flikar_att_exportera:
    raise ValueError("Det finns inga flikar efter README.")

def giltigt_filnamn(namn):
    """Ersätter tecken som inte får användas i Windows-filnamn."""
    return re.sub(r'[<>:"/\\|?*]', "_", namn)

for fliknamn in flikar_att_exportera:
    ws = wb[fliknamn]

    # Letar efter kolumnrubriken i de första 20 raderna
    kolumn_nummer = None
    rubrik_rad = None

    for rad in ws.iter_rows(min_row=1, max_row=min(20, ws.max_row)):
        for cell in rad:
            if cell.value is not None and str(cell.value).strip() == "TXT_format":
                kolumn_nummer = cell.column
                rubrik_rad = cell.row
                break
        if kolumn_nummer is not None:
            break

    if kolumn_nummer is None:
        print(f"Hoppar över '{fliknamn}': Hittar inte kolumnen 'TXT_format'.")
        continue

    # Hämtar alla icke-tomma värden under rubriken
    rader = []

    for rad in range(rubrik_rad + 1, ws.max_row + 1):
        varde = ws.cell(row=rad, column=kolumn_nummer).value

        if varde is not None and str(varde).strip() != "":
            rader.append(str(varde))

    # Sparar i undermappen "TXT filer"
    txt_fil = txt_mapp / f"{giltigt_filnamn(fliknamn)}.txt"

    with open(txt_fil, "w", encoding="utf-8") as fil:
        fil.write("\n".join(rader))

    print(f"Skapade: TXT filer\\{txt_fil.name} ({len(rader)} rader)")

wb.close()
print("Klart.")