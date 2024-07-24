from botcity.maestro import *

from B01_CNPJ import ler_arquivo_cnpj
from B02_Consulta_cnpj_receita import Extracao_guias
from B03_Extrair_pdf import leitura_boleto_certidao
from B04_Planilha_Certidao import planilha_certificado


BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():

    print('ler_arquivo_cnpj')
    ler_arquivo_cnpj()

    print('Extracao_guias')
    Extracao_guias()

    print('leitura_boleto_certidao')
    leitura_boleto_certidao()

    print('planilha_certificado')
    planilha_certificado()


if __name__ == '__main__':
    main()
