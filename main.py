import pandas as pd
import numpy as np
from tkinter.filedialog import askopenfilename


class AnaliseDados:
    def __init__(self, caminho_planilha1, caminho_planilha2):
        self.planilha1 = pd.read_excel(caminho_planilha1)
        self.planilha2 = pd.read_excel(caminho_planilha2)
        self.cpfs_planilha1 = self.planilha1['CPF'].dropna()
        self.matriculas_planilha1 = self.planilha1['Matricula'].dropna()
        self.enderecos_planilha2 = self.planilha2['Endereço'].dropna()
        self.cpfs_planilha2 = self.planilha2['CPF'].dropna()

    def cpfs_comuns(self):
        return self.cpfs_planilha1[self.cpfs_planilha1.isin(self.cpfs_planilha2)]

    def linha_correspondente(self, cpf):
        if cpf in self.cpfs_comuns().values:
            indice_comum = self.cpfs_comuns()[self.cpfs_comuns() == cpf].index.values[0]
            linha_planilha2 = self.planilha2[self.planilha2['CPF'] == cpf]
            linha_planilha1 = self.planilha1.loc[indice_comum].dropna()
            return linha_planilha1, linha_planilha2
        else:
            return None, None
        
if __name__ == "__main__":
    #Ler as planilhas
    caminho_planilha1 = askopenfilename(title = "Selecione a planilha 1", filetypes = [("Excel files", "*.xlsx *.xls")])
    caminho_planilha2 = askopenfilename(title = "Selecione a planilha 2", filetypes = [("Excel files", "*.xlsx *.xls")])

    if not caminho_planilha1 or not caminho_planilha2:
        print("Caminho de arquivo inválido. Por favor, selecione arquivos válidos.")
    else:

        analise = AnaliseDados(caminho_planilha1, caminho_planilha2)
        cpfs_comuns = analise.cpfs_comuns()
        print("CPFs comuns entre as duas planilhas:")
        print(cpfs_comuns)
        if not cpfs_comuns.empty:
            for cpf in cpfs_comuns:
                linha1, linha2 = analise.linha_correspondente(cpf)
                print(f"\nLinha correspondente na Planilha 1 para o CPF {cpf}:")
                print(linha1)
                print(f"\nLinha correspondente na Planilha 2 para o CPF {cpf}:")
                print(linha2)