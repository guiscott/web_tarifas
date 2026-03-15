import pandas as pd
import sqlite3

print("Lendo Excel...")
df = pd.read_excel("C:/Users/guisc/Desktop/web_tarifas/base.xlsx")
print("Linhas carregadas: ", len(df))

print("Criando banco...")
conn = sqlite3.connect("C:/Users/guisc/Desktop/web_tarifas/base.db")
df.to_sql("base", conn, if_exists="replace", index=False)
conn.close()

print("Banco de dados criado com sucesso")