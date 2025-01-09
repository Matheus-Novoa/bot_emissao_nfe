from dados import Dados
from datetime import datetime
from bot_lib import Bot
from config import credentials



def main(dataGeracao, pastaDownload, arqPlanilha, sedes):

    dados = Dados(arqPlanilha)

    if dados.arquivo_progresso.exists():
        with open(dados.arquivo_progresso) as f:
            linha = int(f.read().split()[-1])
        dadosDf = dados.obter_dados()
        dadosDf = dadosDf.iloc[linha:]

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
    bot.bot_setup(pastaDownload)
    
    sede = [texto for texto, var in sedes.items() if var.get()][0]
    bot.entrar(credentials[sede])
    
    bot.definir_data(dataGeracao)
    
    for cliente in dados.itertuples():
        try:
            bot.preencher_campos(cliente, mes, ANO)

            if sede == 'Matriz':
                bot.gerar_nf('Logon do Token', '123456')
            else:
                bot.gerar_nf('Introduzir PIN', '1234')
            
            bot.baixar_nf(dados.arquivo_notas)
        except:
            with open(dados.arquivo_progresso, 'w') as f:
                f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
                raise
        ################### RETORNA E LIMPA OS CAMPOS ###################
        bot.retornar(dados.arquivo_progresso)

    bot.sair()
    bot.fechar_navegador()


if __name__ == '__main__':
    ...