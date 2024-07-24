import aspose.words as aw
import os
import re
import datetime
import pyodbc
import shutil


def leitura_boleto_certidao():

    # CONECTAR BANCO RPA:
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

    conexao, cursor = conexao_bd_rpa("***", "***", "***", "***")
    sql_query = fr"SELECT * FROM DE_CNPJ WHERE STATUS_EXECUCAO = 'CERTIDAO BAIXADA' AND DATA_ABERTURA IS NULL"

    cursor.execute(sql_query)
    DE_CERTIDAO_BASE = cursor.fetchall()

    for linha in DE_CERTIDAO_BASE:

        CNPJ = linha[0]
        print(CNPJ)

        # TRATAR CNPJ
        CNPJ_Tratado = CNPJ.replace('.', '').replace('/', '').replace('-', '')

        arquivo_pasta_pdf = fr'Caminho Diretório'
        arquivo_pasta = fr"{arquivo_pasta_pdf}\{CNPJ_Tratado}.pdf"
        arquivo_pasta_txt = fr"{arquivo_pasta_pdf}\{CNPJ_Tratado}"


        try:

            doc = aw.Document(fr'{arquivo_pasta}')
            doc.save(fr'{arquivo_pasta_txt}.txt')
            doc = (fr'{arquivo_pasta_txt}.txt')

            with open(fr'{arquivo_pasta_txt}.txt', 'r', encoding='charmap', errors='replace') as file:

                linhas = file.readlines()

                print(linhas)

            # Numero de Inscrição:
            for i, linha in enumerate(linhas):

                print(linha)

                if r'NÃMERO DE INSCRIÃÃO' in linha:

                    valorprincipal = linha

                    print(valorprincipal)

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        INSCRICAO = teste.strip().encode('latin1').decode('utf-8')

                        print("Numero de Inscrição:", INSCRICAO)

                    break

            # Data de Abertura
            for i, linha in enumerate(linhas):

                if r'DATA DE ABERTURA' in linha:

                    valorprincipal = linha

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        DATADEABERTURA = teste.strip()

                        print("Data de Abertura:", DATADEABERTURA)

                    break

            # Nome Empresarial
            for i, linha in enumerate(linhas):

                if r'NOME EMPRESARIAL' in linha:

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        NOMEEMPRESARIAL = teste.strip().encode('latin1').decode('utf-8')

                        print("Nome Empresarial:", NOMEEMPRESARIAL)

                    break

            # Nome Fantasia
            for i, linha in enumerate(linhas):

                if r'(NOME DE FANTASIA)' in linha:

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        NOMEFANTASIA = teste.strip()

                        # NOME FANTASIA:
                        print("Nome Fantasia:", NOMEFANTASIA.encode('latin1').decode('utf-8'))

                        # PORTE:
                        PORTE = NOMEFANTASIA.split()[-1]
                        print("PORTE:", PORTE)

                        Titulo_nome_fantasia = fr"{NOMEFANTASIA} {PORTE}"
                        print(Titulo_nome_fantasia)

                    break

            # Codigo e Descrição Principal
            for i, linha in enumerate(linhas):

                if r'PRINCIPAL' in linha:

                    valorprincipal = linha

                    print(valorprincipal)

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        principal = teste.strip().encode('latin1').decode('utf-8')

                        print('PRINCIPAL: ', principal)

                    break

            # Codigo e Descrição Secundárias
            for i, linha in enumerate(linhas):

                if r'SECUNDÃ' in linha:
                    pos = linha.find(r'RIAS')
                    SecundariasPDF = linha[pos:]
                    Secundarias = SecundariasPDF.replace('RIAS', '').replace(' ', '').encode('latin1').decode('utf-8')
                    print('SECUNDÁRIAS: ', Secundarias)

                    break

            # Codigo e Descrição Juridica
            for i, linha in enumerate(linhas):

                if r'DA NATUREZA JURÃ' in linha:

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        Juridica = teste.strip().encode('latin1').decode('utf-8')

                        print('Juridica: ', Juridica)

                    break

            # Lougradouro
            for i, linha in enumerate(linhas):

                if r'LOGRADOURO' in linha:

                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        lougradouro = teste.strip().encode('latin1').decode('utf-8')

                        print('Lougradouro: ', lougradouro)

                    break

            # CEP
            for i, linha in enumerate(linhas):

                if r'CEP' in linha:
                    CEP = linha.encode('latin1').decode('utf-8').replace('CEP BAIRRO/DISTRITO MUNICÍPIO UF', '')
                    print('CEP: ', CEP)

                    break

            # TELEFONE
            for i, linha in enumerate(linhas):

                if r'TELEFONE' in linha:
                    TELEFONE = linha.encode('latin1').decode('utf-8')
                    print('TELEFONE: ', TELEFONE)

                    break

            # FEDERATIVO
            for i, linha in enumerate(linhas):

                if r'FEDERATIVO' in linha:
                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        FEDERATIVO = teste.strip().encode('latin1').decode('utf-8')

                    print('FEDERATIVO: ', FEDERATIVO)

                    break

            # SITUAÇÃO CADASTRAL
            for i, linha in enumerate(linhas):

                if r'CADASTRAL DATA DA SITUAÃ' in linha:
                    # Garante que o índice não saia do limite da lista
                    pos = linha.find('CADASTRAL')
                    SecundariasPDF = linha[pos:]
                    Situacaotxt = SecundariasPDF.encode('latin1').decode('utf-8')
                    Situacao = Situacaotxt.replace('CADASTRAL DATA DA SITUAÇÃO CADASTRAL', '')
                    # print('SITUAÇÃO CADASTRAL: ', Situacao)

                    match = re.search(r'(\w+)\s+(\d{2}/\d{2}/\d{4})', Situacao)
                    status = match.group(1)  # Captura o primeiro grupo de caracteres (\w+)
                    data_extraida = match.group(2)  # Captura o segundo grupo de caracteres (\d{2}/\d{2}/\d{4})
                    print('SITUAÇÃO CADASTRAL:', status)
                    print('DATA SITUAÇÃO CADASTRAL:', data_extraida)

                    break

            # Motivo de situação cadastral
            for i, linha in enumerate(linhas):

                if r'MOTIVO DE SITUAÃ' in linha:
                    # Garante que o índice não saia do limite da lista
                    pos = linha.find('CADASTRAL')
                    SecundariasPDF = linha[pos:]
                    Situacaotxt = SecundariasPDF.encode('latin1').decode('utf-8')
                    MotivoSituacao = Situacaotxt.replace('CADASTRAL DATA DA SITUAÇÃO CADASTRAL', '')
                    print('Motivo de situação cadastral: ', MotivoSituacao)
                    break

            # SITUAÇÃO ESPECIAL
            for i, linha in enumerate(linhas):

                if r'ESPECIAL DATA DA SITUAÃ' in linha:
                    # Garante que o índice não saia do limite da lista
                    if i + 1 < len(linhas):
                        teste = linhas[i + 1].strip()
                        SITUACAO_ESPECIAL = teste.strip().encode('latin1').decode('utf-8')

                    print('SITUAÇÃO ESPECIAL: ', SITUACAO_ESPECIAL)

                    break


            #INSERIR DADOS NO BANCO
            conexao, cursor = conexao_bd_rpa("***", "***", "***", "***")

            sql_query = fr"select IIF(count(*) = 0, 'Não', 'Sim') as TEM_DADOS from DE_CNPJ WHERE STATUS_EXECUCAO = 'CERTIDAO BAIXADA' AND INSCRICAO = '{CNPJ}' "


            print(sql_query)

            cursor.execute(sql_query)
            de_certidao = cursor.fetchall()

            for linha in de_certidao:

                tem_dados = linha[0]

                if tem_dados == 'Sim':

                    print(fr"UPDATE DE_CNPJ SET DATA_ABERTURA = '{DATADEABERTURA}',NOME_EMPRESARIAL = '{NOMEEMPRESARIAL}',TITULO_ESTABELECIMENTO = '{Titulo_nome_fantasia}',COD_DESCRICAO = '{principal}', COD_DESCRICAO_SECUNDARIA = '{Secundarias}', LOUGRADOURO = '{lougradouro}', NUMERO = '{lougradouro}',COMPLEMENTO = '{lougradouro}',CEP = '{CEP}',BAIRRO = '{CEP}',MUNICIPIO = '{CEP}', UF = '{CEP}', ENDERECO = '**', TELEFONE = '**', EFR = '{FEDERATIVO}', SITUACAO = '{status}', DATA_SITUACAO_CADASTRAL = '{data_extraida}', MOTIVO_SITUACAO = '{MotivoSituacao}', SITUACAO_ESPECIAL = '{SITUACAO_ESPECIAL}', DATA_SITUACAO_ESPECIAL = '{SITUACAO_ESPECIAL}' WHERE INSCRICAO = '{CNPJ}'")

                    cursor.execute(
                        fr"UPDATE DE_CNPJ SET DATA_ABERTURA = '{DATADEABERTURA}',NOME_EMPRESARIAL = '{NOMEEMPRESARIAL}',TITULO_ESTABELECIMENTO = '{Titulo_nome_fantasia}',COD_DESCRICAO = '{principal}', COD_DESCRICAO_SECUNDARIA = '{Secundarias}', LOUGRADOURO = '{lougradouro}', NUMERO = '{lougradouro}',COMPLEMENTO = '{lougradouro}',CEP = '{CEP}',BAIRRO = '{CEP}',MUNICIPIO = '{CEP}', UF = '{CEP}', ENDERECO = '{TELEFONE}', TELEFONE = '{TELEFONE}', EFR = '{FEDERATIVO}', SITUACAO = '{status}', DATA_SITUACAO_CADASTRAL = '{data_extraida}', MOTIVO_SITUACAO = '{MotivoSituacao}', SITUACAO_ESPECIAL = '{SITUACAO_ESPECIAL}', DATA_SITUACAO_ESPECIAL = '{SITUACAO_ESPECIAL}' WHERE INSCRICAO = '{CNPJ}'")
                    cursor.commit()

                    break


        except Exception as e:

            print({e})

            cursor.execute(
                fr"UPDATE DE_CNPJ SET STATUS_EXECUCAO = 'VERIFICAR CERTIDAO CNPJ' WHERE INSCRICAO = '{CNPJ}'")
            cursor.commit()

            continue




