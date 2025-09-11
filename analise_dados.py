#Importando as bibliotecas necessárias
import pandas as pd
import numpy as np


#Carregando o dataset das planilhas
planilha1 = pd.read_excel('planilha1.xlsx')
planilha2 = pd.read_excel('planilha2.xlsx')

print("Planilha 1:")
print(planilha1.head())

print("\nPlanilha 2:")
print(planilha2.head())

#Filtrar apenas os CPFs da planilha 1
cpfs_planilha1 = planilha1['CPF'].dropna()#.astype(str).str.zfill(11)

print("\nCPFs na Planilha 1:")
print(cpfs_planilha1)

#Filtrar apenas as matrículas da planilha 1
matriculas_planilha1 = planilha1['Matricula'].dropna()#.astype(str).str.zfill(11)

print("\nMatrículas na Planilha 1:")
print(matriculas_planilha1)

#Filtrar apenas os endereços da planilha 2
enderecos_planilha2 = planilha2['Endereço'].dropna()#.astype(str).str.zfill(11)

print("\nEndereços na Planilha 1:")
print(enderecos_planilha2)

#Filtrar apenas os CPFs da planilha 2
cpfs_planilha2 = planilha2['CPF'].dropna()#.astype(str).str.zfill(11)

print("\nCPFs na Planilha 2:")
print(cpfs_planilha2)

#Verificar se um CPF da planilha 1 está na planilha 2

cpfs_comuns_planilha1_planilha2 = cpfs_planilha1[cpfs_planilha1.isin(cpfs_planilha2)]
print("\nCPFs comuns entre as duas planilhas (Planilha 1 - Planilha 2):")
print(cpfs_comuns_planilha1_planilha2.index.values[0])

#Buscar pelo índice do CPF comum a linha correspondente na planilha 2
if not cpfs_comuns_planilha1_planilha2.empty:
    indice_comum = cpfs_comuns_planilha1_planilha2.index.values[0]
    linha_correspondente_planilha2 = planilha2[planilha2['CPF'] == cpfs_comuns_planilha1_planilha2.iloc[0]]
    linha_correspondente_planilha1 = planilha1.loc[indice_comum].dropna()
    print("\nLinha correspondente na Planilha 2 para o CPF comum:")
    print(linha_correspondente_planilha2)
    print("\nLinha correspondente na Planilha 1 para o CPF comum:")
    print(linha_correspondente_planilha1)