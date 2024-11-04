from botcity.web import WebBot
from botcity.core import DesktopBot
from dados import *
from dotenv import load_dotenv
import sys
sys.path.append(r"C:\Users\novoa\OneDrive\Área de Trabalho\projetos\lib_emissao_nfe")
from bot_lib import Bot


load_dotenv()

if arquivo_progresso.exists():
    with open(arquivo_progresso) as f:
        linha = int(f.read().split()[-1])
    dados = dados.iloc[linha:]

def not_found(label):
    print(f"Element not found: {label}")


def main():

    dataGeracao = '31082024'
    mes = 'AGOSTO'
    ANO = '2024'

    bot = Bot(WebBot, DesktopBot)
    bot.bot_setup(r'C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\NOTA_FICAL_ZN_MAPLE_BEAR\agosto2024')
    bot.entrar()
    bot.definir_data(dataGeracao)
    
    for cliente in dados.itertuples():
        try:
            bot.preencher_campos(cliente, mes, ANO)
            bot.gerar_nf('Logon do Token', '123456')
            bot.baixar_nf(arquivo_notas)
        except:
            with open(arquivo_progresso, 'w') as f:
                f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
                raise
        ################### RETORNA E LIMPA OS CAMPOS ###################
        bot.retornar(arquivo_progresso)

    bot.sair(bot)
    bot.fechar_navegador()


if __name__ == '__main__':
    main()
