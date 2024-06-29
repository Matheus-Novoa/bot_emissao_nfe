import pandas as pd
from pathlib import Path
from tkinter import filedialog


# arquivo_planilha = filedialog.askopenfilename()
arquivo_planilha = Path(r"C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\planilhas\zona_sul\escola_canadenseZS_jun24\Numeração de Boletos_Zona Sul_2024_Junho.xlsx")
arquivo_progresso = arquivo_planilha.parent / 'progresso.log'
arquivo_notas = arquivo_planilha.parent / 'num_notas.txt'
dados = pd.read_excel(arquivo_planilha, 'dados', header=0)

dados['Aluno'] = dados['Aluno'].apply(lambda i: i.split()[0])

dados.loc[dados['Turma'].str.contains('Y1|Year'), 'Acumulador'] = '2'
dados['Acumulador'] = dados['Acumulador'].fillna('1')

dados['Mensalidade'] = dados['Mensalidade'].apply(lambda x: '{:0.2f}'.format(x))
dados['ValorTotal'] = dados['ValorTotal'].apply(lambda x: '{:0.2f}'.format(x))
dados['Alimentação'] = dados['Alimentação'].apply(lambda x: '{:0.2f}'.format(x))


if __name__ == '__main__':
    print(dados)
