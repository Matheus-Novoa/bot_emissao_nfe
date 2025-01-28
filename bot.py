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

    # if dados.arquivo_progresso.exists():
    #     with open(dados.arquivo_progresso) as f:
    #         linha = int(f.read().split()[-1])
    #     dadosDf = dadosDf.iloc[linha:]

    # Identificar o ponto de retomada: a primeira célula não nula
    linha_comeco = dadosDf['Status'].last_valid_index()
    if linha_comeco is None:
        linha_comeco = 0  # Começa do início se não houver valores preenchidos
    else:
        linha_comeco += 1  # Retoma da próxima linha após a última processada
    dadosDf = dadosDf.iloc[linha_comeco:]


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
    
    bot.definir_data(dataGeracao)

    time.sleep(3)
    
    for cliente in dadosDf.itertuples():
        try:
            bot.preencher_campos(cliente, mes, ANO)

            # if sede == 'Matriz':
            #     bot.gerar_nf('Logon do Token', '123456')
            # else:
            #     bot.gerar_nf('Introduzir PIN', '1234')
            
            # bot.baixar_nf(dados.arquivo_notas)

            # Caso o processamento seja bem-sucedido
            dadosDf.at[cliente.Index, 'Status'] = 'OK'
        except:
            dadosDf.at[cliente.Index, 'Status'] = 'ERRO'

            with open(dados.arquivo_progresso, 'w') as f:
                f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
                raise
        finally:
            dados.registra_numero_notas()#num_nota)
        ################### RETORNA E LIMPA OS CAMPOS ###################
        # bot.retornar(dados.arquivo_progresso)
        bot.limpar_campos()

    bot.sair()
    bot.fechar_navegador()


if __name__ == '__main__':
    main('27012025',
         r'C:\Users\novoa\OneDrive\Área de Trabalho\MB_ZS_JAN',
         r"C:\Users\novoa\OneDrive\Área de Trabalho\MB_ZS_JAN\planilha\Numeração de Boletos_Zona Sul_2025_Jan.xlsx",
         {'Matriz': False, 'Zona Sul': True})