import pandas as pd

# Chargement du fichier
df = pd.read_excel("data/Ligue_1.xlsx")

# -------------------------
# Variables créées
# -------------------------

# Efficacité offensive
df["efficacite_offensive"] = (
    df["buts_marques"] / df["xG"]
).round(2)

# Solidité défensive
df["solidite_defensive"] = (
    df["clean_sheets"] / df["journees"]
).round(2)

# Discipline
df["indice_discipline"] = (
    df["cartons_jaunes"] + 3 * df["cartons_rouges"]
)

# Catégorie d'équipe
def categorie_equipe(rang):
    if rang <= 5:
        return "Top 5"
    elif rang <= 12:
        return "Milieu"
    else:
        return "Bas de tableau"

df["categorie_equipe"] = df["rang"].apply(categorie_equipe)

# Sauvegarde
df.to_excel(
    "outputs/Ligue_1_clean.xlsx",
    index=False
)

print("Nettoyage terminé.")