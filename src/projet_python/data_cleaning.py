import pandas as pd
from openpyxl import load_workbook, Workbook
from pathlib import Path


def read_excel_with_openpyxl(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    data = list(ws.values)
    headers = data[0]
    rows = data[1:]

    return pd.DataFrame(rows, columns=headers)


def write_df_with_openpyxl(df, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    ws.append(list(df.columns))

    for row in df.itertuples(index=False):
        ws.append(list(row))

    wb.save(path)


df = read_excel_with_openpyxl("data/Ligue_1.xlsx")


def categorie_equipe(rang):
    if rang <= 5:
        return "Top 5"
    elif rang <= 12:
        return "Milieu"
    else:
        return "Bas de tableau"


df["efficacite_offensive"] = (df["buts_marques"] / df["xG"]).round(2)
df["solidite_defensive"] = (df["clean_sheets"] / df["journees"]).round(2)
df["indice_discipline"] = df["cartons_jaunes"] + 3 * df["cartons_rouges"]
df["categorie_equipe"] = df["rang"].apply(categorie_equipe)

write_df_with_openpyxl(df, "outputs/Ligue_1_clean.xlsx")

print("Nettoyage terminé : outputs/Ligue_1_clean.xlsx")