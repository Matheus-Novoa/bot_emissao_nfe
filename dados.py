import pandas as pd
from pathlib import Path



class Dados:
    def __init__(self, arqPlanilha):
        self.arqPlanilha = Path(arqPlanilha)
        
        self.arquivo_progresso = self.arqPlanilha.parent / 'progresso.log'
        self.arquivo_notas = self.arqPlanilha.parent / 'num_notas.txt'
        
    def obter_dados(self):
        self.dados = pd.read_excel(self.arqPlanilha, 'dados', header=1, skipfooter=1)
        
        self.dados['Aluno'] = self.dados['Aluno'].apply(lambda i: i.split()[0])

        self.dados.loc[self.dados['Turma'].str.contains('Y1|Year'), 'Acumulador'] = '2'
        self.dados['Acumulador'] = self.dados['Acumulador'].fillna('1')

        self.dados['Mensalidade'] = self.dados['Mensalidade'].apply(lambda x: '{:0.2f}'.format(x).replace('.',','))
        self.dados['ValorTotal'] = self.dados['ValorTotal'].apply(lambda x: '{:0.2f}'.format(x).replace('.',','))
        self.dados['Alimentação'] = self.dados['Alimentação'].apply(lambda x: '{:0.2f}'.format(x).replace('.',','))

        return self.dados

if __name__ == '__main__':
    arquivo_planilha = r"C:\Users\novoa\OneDrive\Área de Trabalho\MB_ZS_JAN\planilha\Numeração de Boletos_Zona Sul_2025_Jan.xlsx"
    dados = Dados(arquivo_planilha)
    df = dados.obter_dados()
    print(list(df.itertuples()))