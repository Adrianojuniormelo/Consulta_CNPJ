import csv
import os
import shutil
import pyodbc


def ler_arquivo_cnpj():

    try:
        #CONEXAO BANCO DE DADOS
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

        arquivo_csv = r'Caminho para o diretório'

        # Ler o arquivo CSV
        with open(arquivo_csv, newline='') as csvfile:

            leitor_csv = csv.reader(csvfile)

            for linha in leitor_csv:

                CNPJ = linha[0]

                conexao, cursor = conexao_bd_rpa("***", "***", "***", "***@")
                sql_query = fr"select IIF(count(*) = 0, 'Não', 'Sim') as TEM_DADOS from DE_CNPJ where INSCRICAO = '{CNPJ}'"

                cursor.execute(sql_query)
                de_vendas = cursor.fetchall()

                for linha in de_vendas:

                    tem_dados = linha[0]

                    if tem_dados == 'Não':

                        if CNPJ != 'CNPJs':

                            print(fr"INSERT INTO DE_CNPJ (INSCRICAO) VALUES ( '{CNPJ}')")
                            # INSERIR DADOS NO BANCO DO RPA
                            cursor.execute(fr"INSERT INTO DE_CNPJ (INSCRICAO) VALUES ( '{CNPJ}')")

                            cursor.commit()
                            cursor.close()

        caminho_arquivo_origem = arquivo_csv
        caminho_arquivo_destino = r'Caminho para o diretório'

        if os.path.exists(caminho_arquivo_origem):
            nome_arquivo_destino = os.path.basename(caminho_arquivo_destino)
            novo_caminho_arquivo_destino = os.path.join(os.path.dirname(caminho_arquivo_destino),
                                                        f'CNPJ_OK.csv')

            shutil.move(caminho_arquivo_origem, novo_caminho_arquivo_destino)
        else:
            print(f"O arquivo {caminho_arquivo_origem} não existe.")


    except Exception as e:
        # Trate exceções de conexão aqui
        print(f"Erro ao ler planilha pasta: {e}")


















