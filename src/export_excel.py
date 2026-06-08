import pandas as pd

# =========================
# Chargement des données
# =========================
df = pd.read_excel("outputs/Ligue_1_clean.xlsx")

# =========================
# Création du fichier final
# =========================
output_file = "outputs/dashboard_ligue1.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)

print("Fichier dashboard créé :", output_file)
print("Prochaine étape : ouvrir Excel et créer le Dashboard avec les TCD + segment categorie_equipe.")