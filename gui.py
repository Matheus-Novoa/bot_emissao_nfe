import sys
sys.coinit_flags = 2

import customtkinter as ctk
import json
from bot import main
import warnings
warnings.filterwarnings('ignore')
from dados import Dados


janela = ctk.CTk()
janela.geometry('500x230')
janela.title('Emitir Notas MB')

janela.grid_columnconfigure(0, weight=0, minsize=100)
janela.grid_columnconfigure(1, weight=1)
janela.grid_columnconfigure(2, weight=0)

padx = 5
pady = 5
entryWidth = 250

#------------------------ Data Geração ------------------------
dataGeracaoEntradaTexto = ctk.CTkLabel(janela, text='Data de emissão')
dataGeracaoEntradaTexto.grid(row=0, column=0, padx=padx, pady=pady, sticky="e")

dataGeracaoEntrada = ctk.CTkEntry(janela)
dataGeracaoEntrada.grid(row=0, column=1, padx=padx, pady=pady, sticky="ew")

#------------------------ Pasta download ------------------------
pastaDownloadTexto = ctk.CTkLabel(janela, text='Pasta para Download')
pastaDownloadTexto.grid(row=1, column=0, padx=padx, pady=pady, sticky="e")

pastaDownload = ctk.CTkEntry(janela)
pastaDownload.grid(row=1, column=1, padx=(padx, 0), pady=pady, sticky="ew")

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
botaoDownload.grid(row=1, column=2, padx=(padx, padx), pady=pady, sticky="w")

#------------------------ Arquivo Planilha ------------------------
arqPlanilhaTexto = ctk.CTkLabel(janela, text='Planilha de Dados')
arqPlanilhaTexto.grid(row=2, column=0, padx=padx, pady=pady, sticky="e")

arqPlanilha = ctk.CTkEntry(janela)
arqPlanilha.grid(row=2, column=1, padx=(padx, 0), pady=pady, sticky="ew")

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
botaoPlanilha.grid(row=2, column=2, padx=(padx, padx), pady=pady, sticky="w")

#------------------------ Sede ------------------------
frame_controles = ctk.CTkFrame(janela, fg_color="transparent")
frame_controles.grid(row=3, column=0, columnspan=3, padx=0, pady=pady, sticky="ew")

# Configure columns inside the frame
frame_controles.grid_columnconfigure(0, weight=0)
frame_controles.grid_columnconfigure(1, weight=0)
frame_controles.grid_columnconfigure(2, weight=0)

matriz = ctk.BooleanVar()
zonaSul = ctk.BooleanVar()

matrizCheck = ctk.CTkCheckBox(frame_controles,
                              text='Matriz',
                              variable=matriz,
                              onvalue=True,
                              offvalue=False)


zonaSulCheck = ctk.CTkCheckBox(frame_controles,
                               text='Zona Sul',
                               variable=zonaSul,
                               onvalue=True,
                               offvalue=False)


checkbox_data = {'Matriz': matriz, 'Zona Sul': zonaSul}
#------------------------ Formatar Planilha ------------------------

def botao_formatar_funcao():
    sede = [texto for texto, var in checkbox_data.items() if var.get()][0]
    dados_planilha = Dados(arqPlanilha=arqPlanilha.get(), sede=sede)
    dados_planilha.formata_planilha(arqPlanilha.get())

botao_formatar = ctk.CTkButton(
    frame_controles,
    text='Formatar Planilha',
    width=120,
    command=botao_formatar_funcao
)
matrizCheck.grid(row=0, column=0, padx=(10, 5), pady=0, sticky="w")
zonaSulCheck.grid(row=0, column=1, padx=5, pady=0, sticky="w")
botao_formatar.grid(row=0, column=2, padx=5, pady=0, sticky="w")
    
#------------------------ Iniciar Automação ------------------------
botao = ctk.CTkButton(
    janela,
    text='Iniciar',
    command=lambda: main(dataGeracaoEntrada.get(),
                         pastaDownload.get(),
                         arqPlanilha.get(),
                         checkbox_data
                    )
)

botao.grid(row=4, column=0, columnspan=3, padx=padx, pady=(10, pady), sticky="ew")

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
