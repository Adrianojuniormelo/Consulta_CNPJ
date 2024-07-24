from openpyxl import Workbook
import pyodbc



def planilha_certificado():

    # Função para conectar ao banco de dados
    def conexao_bd_rpa(Server, Database, usuario, senha):
        # CONECTAR BANCO UAU:
        dados_conexao_rpa = pyodbc.connect(
            f"Driver={{SQL Server}};"
            f"Server={Server};"
            f"Database={Database};"
            f"UID={usuario};"
            f"PWD={senha};"
        )

        # VARIÁVEIS DE CONEXÃO DO BD_UAU:
        cursor = dados_conexao_rpa.cursor()

        return dados_conexao_rpa, cursor

    # Criar um novo arquivo Excel
    wb = Workbook()
    ws = wb.active

    # Adicionar os cabeçalhos das colunas
    headers = ["INSCRICAO", "DATA_ABERTURA", "NOME_EMPRESARIAL", "TITULO_ESTABELECIMENTO", "COD_DESCRICAO",
               "COD_DESCRICAO_SECUNDARIA", "LOUGRADOURO", "NUMERO", "COMPLEMENTO", "CEP", "BAIRRO", "MUNICIPIO", "UF",
               "ENDERECO", "TELEFONE", "EFR", "SITUACAO", "DATA_SITUACAO_CADASTRAL", "MOTIVO_SITUACAO", "SITUACAO_ESPECIAL",
               "DATA_SITUACAO_ESPECIAL"]

    ws.append(headers)

    # Conectar ao banco de dados
    conexao, cursor_2 = conexao_bd_rpa("***", "***", "***", "***")

    # Consultar os dados na base de dados
    sql_query = "SELECT * FROM DE_CNPJ WHERE STATUS_EXECUCAO = 'CERTIDAO BAIXADA' AND DATA_ABERTURA IS NOT NULL"
    cursor_2.execute(sql_query)

    # Buscar todas as linhas do resultado da consulta
    tabela_rpa_notas = cursor_2.fetchall()

    # Iterar sobre as linhas e adicionar na planilha
    for linha in tabela_rpa_notas:
        ws.append(list(linha))

    # Salvar o arquivo Excel
    wb.save("consulta_cnpj.xlsx")

    # Fechar a conexão com o banco de dados
    conexao.close()



