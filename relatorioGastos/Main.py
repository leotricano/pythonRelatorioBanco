import pandas as pd
from datetime import datetime

# Função para adicionar transação na planilha
def adicionar_transacao(valor, metodo_pagamento, descricao):

    # Captura a data e hora atuais
    data_hora_atual = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Carregar ou criar planilha
    try:
        df = pd.read_excel('Compras.xlsx')
    except FileNotFoundError: #Se o arquivo não existir, cria um DataFrame vazio com as colunas: Data, Valor, Método de Pagamento, e Descrição.

        # Inicializando o DataFrame com tipos de dados especificados para evitar o aviso
        df = pd.DataFrame({
            'Data e Hora': pd.Series(dtype='datetime64[ns]'),
            'Valor': pd.Series(dtype='float'),
            'Método de Pagamento': pd.Series(dtype='str'),
            'Descrição': pd.Series(dtype='str')
        })

    nova_transacao = {'Data e Hora': pd.to_datetime(data_hora_atual, format='%d/%m/%Y %H:%M:%S'),
                      'Valor': float(f'{valor:.2f}'),
                      'Método de Pagamento':metodo_pagamento,
                      'Descrição': descricao }

    # Adiciona a nova transação ao DataFrame
    df = pd.concat( [df,pd.DataFrame([nova_transacao])],ignore_index=True)

    # Usar o XlsxWriter para formatar a coluna Valor como moeda (R$)
    writer = pd.ExcelWriter('Compras.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False)

    #Salvar e voltar a planilha
    df.to_excel('Compras.xlsx', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    format_currency = workbook.add_format({'num_format':'R$ #,##0.00'})
    worksheet.set_column('B:B', None, format_currency)

    writer._save()


#Testar =====================================================
adicionar_transacao ( 200.00, 'Pix', 'Japones')