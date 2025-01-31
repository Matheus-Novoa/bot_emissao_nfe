import pandas as pd
from pathlib import Path
from openpyxl import load_workbook



class Dados:
    def __init__(self, arqPlanilha):
        self.arqPlanilha = Path(arqPlanilha)
        
        self.arquivo_progresso = self.arqPlanilha.parent / 'progresso.log'
        self.arquivo_notas = self.arqPlanilha.parent / 'num_notas.txt'
        
    def obter_dados(self):
        self.dados = pd.read_excel(self.arqPlanilha, 'dados', header=1, skipfooter=1)

        if 'Notas' not in self.dados.columns:
            self.dados['Notas'] = None
        
        self.dados['Aluno'] = self.dados['Aluno'].apply(lambda i: i.split()[0])

        self.dados.loc[self.dados['Turma'].str.contains('Y1|Year'), 'Acumulador'] = '2'
        self.dados['Acumulador'] = self.dados['Acumulador'].fillna('1')

        self.dados['Mensalidade'] = self.dados['Mensalidade'].apply(lambda x: '{:0.2f}'.format(x).replace('.',','))
        self.dados['ValorTotal'] = self.dados['ValorTotal'].apply(lambda x: '{:0.2f}'.format(x).replace('.',','))
        self.dados['Alimentação'] = self.dados['Alimentação'].apply(lambda x: '{:0.2f}'.format(x).replace('.',','))

        return self.dados
    

    def registra_numero_notas(self, index_df, num_nota):
        wb = load_workbook(self.arqPlanilha)
        sheet = wb['dados']

        # Adicionar a coluna 'Status' se não existir
        if 'Notas' not in [celula.value for celula in sheet[2]]:
            sheet.cell(row=2, column=sheet.max_column + 1).value = 'Notas'

        # # Adicionar a coluna 'Status' se não existir
        # if 'Status' not in [celula.value for celula in sheet[2]]:
        #     sheet.cell(row=2, column=sheet.max_column + 1).value = 'Status'

        # Atualizar os valores da coluna 'Status' a partir do ponto processado
        # status_col_index = [celula.value for celula in sheet[2]].index('Status') + 1
        notas_col_index = [celula.value for celula in sheet[2]].index('Notas') + 1

        # for i, num_nota in enumerate(self.dados['Notas'], start=3):
        sheet.cell(row=index_df+3, column=notas_col_index).value = num_nota

        wb.save(self.arqPlanilha)



if __name__ == '__main__':
    arquivo_planilha = r"C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\planilhas\zona_norte\escola_canadenseZN_dez24\Maple Bear Dez 24.xlsx"
    dados = Dados(arquivo_planilha)
    df = dados.obter_dados()
    # print(df)
    print(list(df[df['Notas'].isna()].itertuples()))
    # print(list(df.itertuples()))