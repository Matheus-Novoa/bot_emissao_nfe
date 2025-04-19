from dados import Dados
from datetime import datetime
from bot_lib import Bot
from config import credentials
from dotenv import load_dotenv
import os
import time



def main(dataGeracao, pastaDownload, arqPlanilha, sedes):

    load_dotenv()

    sede = [texto for texto, var in sedes.items() if var.get()][0]
    # sede = [texto for texto, var in sedes.items() if var][0] # apenas para teste

    dados = Dados(arqPlanilha, sede)

    # dados.formata_planilha()
    dadosDf = dados.obter_dados()
    if not dadosDf:
        raise('Planilha original não formatada. Realize a formatação antes de executar o programa.')

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
    
    try:
        bot.entrar(credentials[sede])
    except:
        bot.sair()
        bot.entrar(credentials[sede])
    
    bot.definir_data(dataGeracao)

    time.sleep(3)
    
    for cliente in df_afazer.itertuples():
        try:
            status = bot.preencher_campos(cliente, mes, ANO)
            if not status:
                continue

            if sede == 'Matriz':
                bot.gerar_nf('Logon do Token', '123456')
            else:
                bot.gerar_nf('Introduzir PIN', '1234')
            
            num_nota = bot.baixar_nf(dados.arquivo_notas)

            # Caso o processamento seja bem-sucedido
            df_afazer.at[cliente.Index, 'Notas'] = num_nota

            dados.registra_numero_notas(cliente.Index, num_nota)
            bot.retornar()
        except Exception as error:
            print(error)
            with open(dados.arquivo_progresso, 'w') as f:
                f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
                raise
        finally:
            ################### RETORNA E LIMPA OS CAMPOS ###################
            bot.limpar_campos()

    bot.sair()
    bot.fechar_navegador()


if __name__ == '__main__':
    main('28022025',
         r'C:/Users/novoa/OneDrive/Área de Trabalho/notas_MB/NOTA_FISCAL_ZS_MAPLE_BEAR/FEVEREIRO2025',
         r"C:/Users/novoa/OneDrive/Área de Trabalho/notas_MB/planilhas/zona_sul/escola_canadenseZS_fev25/Boletos_Zona Sul_2025_Fevereiro.xlsx",
         {'Matriz': False, 'Zona Sul': True})