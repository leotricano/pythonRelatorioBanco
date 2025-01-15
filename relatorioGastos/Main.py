import pandas as pd #Excell
import tabula
import locale
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from datetime import datetime

# Oculta a janela principal do Tkinter
Tk().withdraw()

file_path = askopenfilename(title="Selecione o arquivo",
                            filetype=[('Arquivo PDF', "*.PDF"),
                                      ("Todos os arquivos", "*.*")])
if file_path:
    try:
        # Lê tabelas do PDF
        input_file = tabula.read_pdf(file_path, pages="all")

        # Criando um nome de arquivo de saída com um padrão
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        # Data e hora atual
        output_file = f"tabela_extraida_{current_time}.csv"

        # Salvar dados do PDF em CSV
        tabula.convert_into(file_path, output_file, output_format="csv", pages="all")

        # Ler os dados em DataFrame
        df = pd.read_csv(output_file, header=None)
        print(df)

        if len(df.columns) >= 2:
            df = df.dropna(axis=1, how='all')
            df.columns = ["Data", "Local", "Valor"]
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
            df['Data'] = pd.to_datetime('2025 ' + df['Data'], format='%Y %d %b', errors='coerce')
            '''# Filtrando linhas que contenham palavras-chave específicas
            keywords = ["Pagamento", "Saldo", "Crédito"]
            pattern = '|'.join(keywords)
            print(df)
            df = df[~df['Local'].str.contains(pattern, flags=re.IGNORECASE, na=False)]'''
            print(df)

            # Limpeza e filtragem da coluna 'Valor'
            try:
                # Remove o "R$", substitui vírgulas por pontos e remove espaços
                df['Valor'] = (
                    df['Valor']
                    .str.replace("R$", "", regex=False)  # Remove "R$"
                    .str.replace(",", ".", regex=False)  # Substitui vírgulas por pontos
                    .str.strip()  # Remove espaços extras
                )

                # Converte para float antes de fazer a filtragem de valores negativos
                df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')

                # Remove valores negativos
                df = df[df['Valor'] >= 0]  # Filtra valores maiores ou iguais a 0
                print(df)
                # Calcula a soma
                soma_valor = df["Valor"].sum()
                print(f"Soma dos valores: {soma_valor}")
            except Exception as e:
                print(f"Erro ao processar a coluna 'Valor': {e}")
                soma_valor = 0  # Define um valor padrão em caso de erro

            # Calcular o resultado final
            try:
                salario = float(input("Digite seu salário: "))
                resultado = salario - soma_valor
                print(f"Resultado após subtração: {resultado}")
                df.to_excel(f"tabela_extraida_{current_time}.xlsx", index=False)
            except ValueError:
                print("Entrada inválida para salário. Digite um número válido.")

        else:
            print(f"O arquivo gerado possui apenas {len(df.columns)} colunas. Verifique o conteúdo do PDF.")
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
else:
    print("Nenhum arquivo foi selecionado.")













