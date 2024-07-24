import os

import pygetwindow as gw
import time
import pyautogui
import pyodbc
from botcity.web import WebBot, Browser, By
import subprocess
from pynput.keyboard import Controller



def Extracao_guias():


    bot = WebBot()
    bot.headless = False
    keyboard = Controller()

    # ABRIR UAU FISCAL
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
    # LER BASE NFTS
    sql_query = fr"SELECT * FROM DE_CNPJ WHERE INSCRICAO = '17.516.113/0001-47'"

    cursor.execute(sql_query)
    DE_IPTU_BASE = cursor.fetchall()

    tentativas = 0
    max_tentativas = 1

    while tentativas < max_tentativas:

        for linha in DE_IPTU_BASE:

            try:

                print('### INICIO DO PROCESSO ###')

                CNPJ = linha[0]

                caminho_executavel_chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
                url = 'https://solucoes.receita.fazenda.gov.br/Servicos/cnpjreva/cnpjreva_solicitacao.asp'

                # Use subprocess.Popen para abrir a URL em uma nova janela do Chrome
                subprocess.Popen([caminho_executavel_chrome, url])

                time.sleep(5)

                # CLICAR CNPJ
                print('CNPJ')
                center = pyautogui.locateCenterOnScreen('CNPJ.png')
                if center:
                    pyautogui.click(231,  568)
                    time.sleep(2)
                    pyautogui.write(CNPJ)


                #QUEBRAR CAPTCHA
                time.sleep(10)

                #CLICAR EM CONSULTAR
                print('CONSULTA')
                center = pyautogui.locateCenterOnScreen('CNPJ.png')
                if center:
                 pyautogui.click( 149,  643)

                time.sleep(5)

                #BAIXAR GUIA CERTIDÃO
                print('CERTIDÃO')
                center = pyautogui.locateCenterOnScreen('CERTIDAO.png')
                if center:
                    pyautogui.hotkey('ctrl', 'p')

                time.sleep(3)

                #SALVAR
                center = pyautogui.locateCenterOnScreen('PDF.png')
                if center:
                    pyautogui.press('enter')

                time.sleep(5)

                #SALVAR COMO JANELA:
                janela_alvo = "Salvar como"
                janelas_abertas = gw.getWindowsWithTitle(janela_alvo)

                #TRATAR CNPJ
                CNPJ_Tratado = CNPJ.replace('.','').replace('/','').replace('-','')

                print(CNPJ_Tratado)

                #CRIAÇÃO DA PASTA
                caminho_pasta_completo = fr"Caminho para o diretório"
                os.makedirs(caminho_pasta_completo, exist_ok=True)

                #Nome Arquivo
                nome_arquivo = fr"{caminho_pasta_completo}\{CNPJ_Tratado}.pdf"

                # Verifique se a janela alvo está aberta
                if janelas_abertas:

                    keyboard.type(nome_arquivo + '\n')

                    time.sleep(6)

                    print('SALVAR NOVAMENTE')
                    janela_alvo = "Confirmar Salvar como"
                    janelas_abertas = gw.getWindowsWithTitle(janela_alvo)

                    if janelas_abertas:

                        pyautogui.press('tab')
                        time.sleep(2)
                        pyautogui.press('enter')

                    #ATUALIZAR INFORMAÇÃO NO BANCO DE DADOS

                    time.sleep(3)

                    cursor.execute(
                        fr"UPDATE DE_CNPJ SET STATUS_EXECUCAO = 'CERTIDAO BAIXADA' WHERE INSCRICAO = '{CNPJ}'")
                    cursor.commit()

                    #FECHAR NAVEGADOR
                    os.system("taskkill /F /IM chrome.exe")

                    print('### TÉRMINO DO PROCESSO ###')


            except Exception as e:

                # Trate exceções de conexão aqui
                print(f" FINALIZADO COM ERRO: {e}")

                cursor.execute(
                    fr"UPDATE DE_CNPJ SET STATUS_EXECUCAO = 'VERIFICAR CERTIDAO CNPJ' WHERE INSCRICAO = '{CNPJ}'")
                cursor.commit()

                continue




