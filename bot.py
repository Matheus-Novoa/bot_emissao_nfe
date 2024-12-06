from dados import *
from dotenv import load_dotenv
import customtkinter as ctk
from datetime import datetime
import json
import sys
sys.path.append(r"C:\Users\novoa\OneDrive\Área de Trabalho\projetos\lib_emissao_nfe")
from bot_lib import Bot


load_dotenv()

if arquivo_progresso.exists():
    with open(arquivo_progresso) as f:
        linha = int(f.read().split()[-1])
    dados = dados.iloc[linha:]


def main(dataGeracao):

    data = datetime.strptime(dataGeracao.get(), '%d%m%Y')
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
    bot.bot_setup(r'C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\NOTA_FICAL_ZN_MAPLE_BEAR\agosto2024')
    bot.entrar()
    bot.definir_data(dataGeracao.get())
    
    # for cliente in dados.itertuples():
    #     try:
    #         bot.preencher_campos(cliente, mes, ANO)
    #         bot.gerar_nf('Logon do Token', '123456')
    #         bot.baixar_nf(arquivo_notas)
    #     except:
    #         with open(arquivo_progresso, 'w') as f:
    #             f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
    #             raise
    #     ################### RETORNA E LIMPA OS CAMPOS ###################
    #     bot.retornar(arquivo_progresso)

    bot.sair()
    bot.fechar_navegador()


if __name__ == '__main__':
    arquivoCache = '.cache'

    janela = ctk.CTk()
    janela.geometry('500x200')
    janela.title('Emitir Notas MB')

    padx = 10
    pady= 5
    entryWidth = 180

    #------------------------ Data Geração ------------------------
    dataGeracaoEntrada = ctk.CTkEntry(janela, placeholder_text='Data de emissão')
    dataGeracaoEntrada.pack()

    botao = ctk.CTkButton(janela, text='Iniciar', command=lambda: main(dataGeracaoEntrada))
    botao.pack()

    try:
        with open(arquivoCache) as f:
            dados = json.load(f)
            dataGeracaoEntrada.delete(0, ctk.END)
            dataGeracaoEntrada.insert(0, dados['data'])
    except FileNotFoundError:
        ...
    except json.decoder.JSONDecodeError:
        ...
    
    def gravar_cache():
        with open(arquivoCache, 'w') as f:
            dados = {
                'data': dataGeracaoEntrada.get()
            }
            json.dump(dados, f)
        janela.destroy()

    janela.protocol('WM_DELETE_WINDOW', gravar_cache)

    janela.mainloop()
