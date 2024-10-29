from datetime import datetime
from sys import displayhook
import pandas as pd #Excell
import tabula

# Testando leitura de tabela no PDF
listas_tabelas = tabula.read_pdf("Nubank_2023-09-04.pdf", pages="all")
print(len(listas_tabelas))

#Mostrando dados coletados no console
for tabela in listas_tabelas:
    displayhook(tabela)

    #Convertendo de PDF para CSV
tabula.convert_into("Nubank_2023-09-04.pdf", "Nubank.csv", output_format="csv", pages="all")

df = pd.read_csv("Nubank.csv", names=["Data":pd.Series(datetime), "Em branco", "Descrição", "Valor"])

# Removendo a coluna "Em branco"
df = df.drop(columns=["Em branco"])

# Usar o XlsxWriter para formatar a coluna Valor como moeda (R$)

writer = pd.ExcelWriter('Nubank.xlsx', engine='xlsxwriter')
df.to_excel(writer, index=False, sheet_name='Sheet1')
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Configurar a formatação para a coluna de valores em formato de moeda
format_currency = workbook.add_format({'num_format': 'R$ #,##0.00'})
worksheet.set_column('C:C', None, format_currency)  # Ajuste o índice conforme a coluna correta

writer._save()

print("Arquivo 'Nubank.xlsx' criado com sucesso.")


#Testar ===========
