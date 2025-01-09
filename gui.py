import sys
sys.coinit_flags = 2

import customtkinter as ctk
import json
from bot import main
import warnings
warnings.filterwarnings('ignore')


janela = ctk.CTk()
janela.geometry('450x200')
janela.title('Emitir Notas MB')

padx = 10
pady= 5
entryWidth = 250

#------------------------ Data Geração ------------------------
dataGeracaoEntradaTexto = ctk.CTkLabel(janela, text='Data de emissão')
dataGeracaoEntradaTexto.grid(row=0, column=0, pady=pady)

dataGeracaoEntrada = ctk.CTkEntry(janela)
dataGeracaoEntrada.grid(row=0, column=1, padx=padx, pady=pady)

#------------------------ Pasta download ------------------------
pastaDownloadTexto = ctk.CTkLabel(janela, text='Pasta para Download')
pastaDownloadTexto.grid(row=1, column=0, padx=padx, pady=pady)

pastaDownload = ctk.CTkEntry(janela, width=entryWidth)
pastaDownload.grid(row=1, column=1, padx=padx, pady=pady)

def selecionar_diretorio(entrada):
    caminho = ctk.filedialog.askdirectory(title="Selecione uma pasta")
    if caminho:
        entrada.delete(0, ctk.END)  # Limpa o campo de entrada
        entrada.insert(0, caminho)  # Insere o caminho do diretório

botaoDownload = ctk.CTkButton(
    janela,
    text='...',
    width=15,
    fg_color='gray',
    command=lambda e=pastaDownload: selecionar_diretorio(e)
)
botaoDownload.grid(row=1, column=2, padx=2, pady=pady)

#------------------------ Arquivo Planilha ------------------------
arqPlanilhaTexto = ctk.CTkLabel(janela, text='Planilha de Dados')
arqPlanilhaTexto.grid(row=2, column=0, padx=padx, pady=pady)

arqPlanilha = ctk.CTkEntry(janela, width=entryWidth)
arqPlanilha.grid(row=2, column=1, padx=padx, pady=pady)

def selecionar_arquivo(entrada):
    caminho = ctk.filedialog.askopenfilename(title="Selecione um arquivo")
    if caminho:
        entrada.delete(0, ctk.END)  # Limpa o campo de entrada
        entrada.insert(0, caminho)  # Insere o caminho do arquivo

botaoPlanilha = ctk.CTkButton(
    janela,
    text='...',
    width=15,
    fg_color='gray',
    command=lambda e=arqPlanilha: selecionar_arquivo(e)
)
botaoPlanilha.grid(row=2, column=2, padx=2, pady=pady)

#------------------------ Sede ------------------------
matriz = ctk.BooleanVar()
zonaSul = ctk.BooleanVar()

matrizCheck = ctk.CTkCheckBox(janela,
                              text='Matriz',
                              variable=matriz,
                              onvalue=True,
                              offvalue=False)

zonaSulCheck = ctk.CTkCheckBox(janela,
                               text='Zona Sul',
                               variable=zonaSul,
                               onvalue=True,
                               offvalue=False)

matrizCheck.grid(row=3, column=0, padx=padx, pady=pady)
zonaSulCheck.grid(row=3, column=1, padx=padx, pady=pady)

checkbox_data = {'Matriz': matriz, 'Zona Sul': zonaSul}

#------------------------ Iniciar Automação ------------------------
botao = ctk.CTkButton(
    janela,
    text='Iniciar',
    width=entryWidth,
    command=lambda: main(dataGeracaoEntrada.get(),
                         pastaDownload.get(),
                         arqPlanilha.get(),
                         checkbox_data
                    )
)

botao.grid(row=4, column=0, columnspan=3, padx=padx, pady=pady)

#------------------------ Carrega cache ------------------------
arquivoCache = '.cache'
try:
    with open(arquivoCache) as f:
        dados = json.load(f)
        dataGeracaoEntrada.delete(0, ctk.END)
        dataGeracaoEntrada.insert(0, dados['data'])

        pastaDownload.delete(0, ctk.END)
        pastaDownload.insert(0, dados['download'])

        arqPlanilha.delete(0, ctk.END)
        arqPlanilha.insert(0, dados['planilha'])
except FileNotFoundError:
    ...
except json.decoder.JSONDecodeError:
    ...

#------------------------ Grava cache ------------------------
def gravar_cache():
    with open(arquivoCache, 'w') as f:
        dados = {
            'data': dataGeracaoEntrada.get(),
            'download': pastaDownload.get(),
            'planilha': arqPlanilha.get(),
        }
        json.dump(dados, f)
    janela.destroy()

janela.protocol('WM_DELETE_WINDOW', gravar_cache)

janela.mainloop()
