from dados import Dados
from datetime import datetime
from bot_lib import Bot
from config import credentials
from dotenv import load_dotenv
import os
import time



def main(dataGeracao, pastaDownload, arqPlanilha, sedes):

    load_dotenv()

    dados = Dados(arqPlanilha)
    dadosDf = dados.obter_dados()

    df_afazer = dadosDf[dadosDf['Notas'].isna()]

    data = datetime.strptime(dataGeracao, '%d%m%Y')
    meses = {
    '1': "Janeiro",
    '2': "Fevereiro",
    '3': "Março",
    '4': "Abril",
    '5': "Maio",
    '6': "Junho",
    '7': "Julho",
    '8': "Agosto",
    '9': "Setembro",
    '10': "Outubro",
    '11': "Novembro",
    '12': "Dezembro"
    }
    mes = meses[str(data.month)]
    ANO = str(data.year)

    bot = Bot()
    bot.bot_setup(pastaDownload, os.getenv('PROFILE_PATH'))
    
    # sede = [texto for texto, var in sedes.items() if var.get()][0]
    sede = [texto for texto, var in sedes.items() if var][0] # apenas para teste
    try:
        bot.entrar(credentials[sede])
    except:
        bot.sair()
        bot.entrar(credentials[sede])
    
    bot.definir_data(dataGeracao)

    time.sleep(3)
    
    for cliente in df_afazer.itertuples():
        try:
            bot.preencher_campos(cliente, mes, ANO)

            if sede == 'Matriz':
                bot.gerar_nf('Logon do Token', '123456')
            else:
                bot.gerar_nf('Introduzir PIN', '1234')
            
            num_nota = bot.baixar_nf(dados.arquivo_notas)

            # Caso o processamento seja bem-sucedido
            df_afazer.at[cliente.Index, 'Notas'] = num_nota
        except:
            # dadosDf.at[cliente.Index, 'Status'] = 'ERRO'

            with open(dados.arquivo_progresso, 'w') as f:
                f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
                raise
        finally:
            dados.registra_numero_notas(cliente.Index, num_nota)
        ################### RETORNA E LIMPA OS CAMPOS ###################
        bot.retornar()
        bot.limpar_campos()

    bot.sair()
    bot.fechar_navegador()


if __name__ == '__main__':
    main('31122024',
         r'C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\NOTA_FICAL_ZN_MAPLE_BEAR\dezembro2024',
         r"C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\planilhas\zona_norte\escola_canadenseZN_dez24\Maple Bear Dez 24.xlsx",
         {'Matriz': True, 'Zona Sul': False})