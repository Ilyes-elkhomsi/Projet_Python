import pandas as pd

with pd.ExcelWriter("outputs/dashboard.xlsx") as writer:
    
    df_classement.to_excel(writer, sheet_name="classement")
    df_attack.to_excel(writer, sheet_name="attaque")
    df_defense.to_excel(writer, sheet_name="defense")
    df_home_away.to_excel(writer, sheet_name="home_away")
    df_players.to_excel(writer, sheet_name="joueurs")
    df_kpis.to_excel(writer, sheet_name="kpis")