import subprocess

required_libraries = ['pandas', 'shutil', 'xlsxwriter']

for library in required_libraries:
    try:
        __import__(library)
    except ImportError:
        subprocess.check_call(['pip', 'install', library])

import os
import pandas as pd
from shutil import copyfile
from datetime import datetime


def gera_notificacao(csv_data):
    #REORGANIZA O CSV PARA MONTAR UMA NOTIFICAÇÃO DE LANÇAMENTO COM OS DADOS DAQUI
    # ACREDITO QUE ESSA REORGANIZAÇÃO NÃO SEJA NECESSÁRIA PARA A EMISSÃO DA NOTIFICAÇÃO VIA STM DIRETO
    header_output = ['Tipo de Registro',
                     #'Código do Objeto Cliente',
                     #'Número do Lote',
                     #'Cartão de Postagem',
                     #'Número do Contrato',
                     #'Serviço Adicional',
                     #'Identificador do Arquivo Spool',
                     #'Nome do Arquivo Spool',
                     #'Identificador do Arquivo Complementar',
                     'Nome do Arquivo Complementar',
                     'Logradouro',
                     'Número',
                     'Complemento',
                     'Bairro',
                     'Cidade',
                     'Estado',
                     'CEP',
                     'Modelo do Layout',
                     'Matrícula',
                     'CPF/CNPJ do Contribuinte',
                     'Contribuinte',
                     'Número do Processo',
                     'Usuário',
                     'Matrícula_Usuário',
                     'Observação', 

                    'E_1-Exercício',
                    'E_1-Tipo do Imóvel - Anterior',
                    'E_1-Área Terreno - Anterior',
                    'E_1-Uso do Imóvel - Anterior',
                    'E_1-Área Construída Unidade - Anterior',
                    'E_1-Área Total Edificada - Anterior',
                    'E_1-Fração Ideal - Anterior',
                    'E_1-Situação da Quadra - Anterior',
                    'E_1-Topografia - Anterior',
                    'E_1-Pedologia - Anterior',
                    'E_1-Tipo de Construção - Anterior',
                    'E_1-Alinhamento - Anterior',
                    'E_1-Situação da Edificação - Anterior',
                    'E_1-Situação da Unidade - Anterior',
                    'E_1-Estrutura da Construção - Anterior',
                    'E_1-Cobertura - Anterior',
                    'E_1-Paredes - Anterior',
                    'E_1-Revestimento Fachada - Anterior',
                    #'E_1-Valor Venal da Edificação - Anterior',
                    #'E_1-Valor Venal do Terreno - Anterior',
                    'E_1-Valor Venal do Imóvel - Anterior',
                    'E_1-Valor do Imposto - Anterior',

                    'E_1-Tipo do Imóvel - Atual',
                    'E_1-Área Terreno - Atual',
                    'E_1-Uso do Imóvel - Atual',
                    'E_1-Área Construída Unidade - Atual',
                    'E_1-Área Total Edificada - Atual',
                    'E_1-Fração Ideal - Atual',
                    'E_1-Situação da Quadra - Atual',
                    'E_1-Topografia - Atual',
                    'E_1-Pedologia - Atual',
                    'E_1-Tipo de Construção - Atual',
                    'E_1-Alinhamento - Atual',
                    'E_1-Situação da Edificação - Atual',
                    'E_1-Situação da Unidade - Atual',
                    'E_1-Estrutura da Construção - Atual',
                    'E_1-Cobertura - Atual',
                    'E_1-Paredes - Atual',
                    'E_1-Revestimento Fachada - Atual',
                    #'E_1-Valor Venal da Edificação - Atual',
                    #'E_1-Valor Venal do Terreno - Atual',
                    'E_1-Valor Venal do Imóvel - Atual',
                    'E_1-Valor do Imposto - Atual',
                    #'E_1-Diferença',

                    'E_2-Exercício',
                    'E_2-Tipo do Imóvel - Anterior',
                    'E_2-Área Terreno - Anterior',
                    'E_2-Uso do Imóvel - Anterior',
                    'E_2-Área Construída Unidade - Anterior',
                    'E_2-Área Total Edificada - Anterior',
                    'E_2-Fração Ideal - Anterior',
                    'E_2-Situação da Quadra - Anterior',
                    'E_2-Topografia - Anterior',
                    'E_2-Pedologia - Anterior',
                    'E_2-Tipo de Construção - Anterior',
                    'E_2-Alinhamento - Anterior',
                    'E_2-Situação da Edificação - Anterior',
                    'E_2-Situação da Unidade - Anterior',
                    'E_2-Estrutura da Construção - Anterior',
                    'E_2-Cobertura - Anterior',
                    'E_2-Paredes - Anterior',
                    'E_2-Revestimento Fachada - Anterior',
                    #'E_2-Valor Venal da Edificação - Anterior',
                    #'E_2-Valor Venal do Terreno - Anterior',
                    'E_2-Valor Venal do Imóvel - Anterior',
                    'E_2-Valor do Imposto - Anterior',

                    'E_2-Tipo do Imóvel - Atual',
                    'E_2-Área Terreno - Atual',
                    'E_2-Uso do Imóvel - Atual',
                    'E_2-Área Construída Unidade - Atual',
                    'E_2-Área Total Edificada - Atual',
                    'E_2-Fração Ideal - Atual',
                    'E_2-Situação da Quadra - Atual',
                    'E_2-Topografia - Atual',
                    'E_2-Pedologia - Atual',
                    'E_2-Tipo de Construção - Atual',
                    'E_2-Alinhamento - Atual',
                    'E_2-Situação da Edificação - Atual',
                    'E_2-Situação da Unidade - Atual',
                    'E_2-Estrutura da Construção - Atual',
                    'E_2-Cobertura - Atual',
                    'E_2-Paredes - Atual',
                    'E_2-Revestimento Fachada - Atual',
                    #'E_2-Valor Venal da Edificação - Atual',
                    #'E_2-Valor Venal do Terreno - Atual',
                    'E_2-Valor Venal do Imóvel - Atual',
                    'E_2-Valor do Imposto - Atual',
                    #'E_2-Diferença',

                    'E_3-Exercício',
                    'E_3-Tipo do Imóvel - Anterior',
                    'E_3-Área Terreno - Anterior',
                    'E_3-Uso do Imóvel - Anterior',
                    'E_3-Área Construída Unidade - Anterior',
                    'E_3-Área Total Edificada - Anterior',
                    'E_3-Fração Ideal - Anterior',
                    'E_3-Situação da Quadra - Anterior',
                    'E_3-Topografia - Anterior',
                    'E_3-Pedologia - Anterior',
                    'E_3-Tipo de Construção - Anterior',
                    'E_3-Alinhamento - Anterior',
                    'E_3-Situação da Edificação - Anterior',
                    'E_3-Situação da Unidade - Anterior',
                    'E_3-Estrutura da Construção - Anterior',
                    'E_3-Cobertura - Anterior',
                    'E_3-Paredes - Anterior',
                    'E_3-Revestimento Fachada - Anterior',
                    #'E_3-Valor Venal da Edificação - Anterior',
                    #'E_3-Valor Venal do Terreno - Anterior',
                    'E_3-Valor Venal do Imóvel - Anterior',
                    'E_3-Valor do Imposto - Anterior',

                    'E_3-Tipo do Imóvel - Atual',
                    'E_3-Área Terreno - Atual',
                    'E_3-Uso do Imóvel - Atual',
                    'E_3-Área Construída Unidade - Atual',
                    'E_3-Área Total Edificada - Atual',
                    'E_3-Fração Ideal - Atual',
                    'E_3-Situação da Quadra - Atual',
                    'E_3-Topografia - Atual',
                    'E_3-Pedologia - Atual',
                    'E_3-Tipo de Construção - Atual',
                    'E_3-Alinhamento - Atual',
                    'E_3-Situação da Edificação - Atual',
                    'E_3-Situação da Unidade - Atual',
                    'E_3-Estrutura da Construção - Atual',
                    'E_3-Cobertura - Atual',
                    'E_3-Paredes - Atual',
                    'E_3-Revestimento Fachada - Atual',
                    #'E_3-Valor Venal da Edificação - Atual',
                    #'E_3-Valor Venal do Terreno - Atual',
                    'E_3-Valor Venal do Imóvel - Atual',
                    'E_3-Valor do Imposto - Atual',
                    #'E_3-Diferença',

                    'E_4-Exercício',
                    'E_4-Tipo do Imóvel - Anterior',
                    'E_4-Área Terreno - Anterior',
                    'E_4-Uso do Imóvel - Anterior',
                    'E_4-Área Construída Unidade - Anterior',
                    'E_4-Área Total Edificada - Anterior',
                    'E_4-Fração Ideal - Anterior',
                    'E_4-Situação da Quadra - Anterior',
                    'E_4-Topografia - Anterior',
                    'E_4-Pedologia - Anterior',
                    'E_4-Tipo de Construção - Anterior',
                    'E_4-Alinhamento - Anterior',
                    'E_4-Situação da Edificação - Anterior',
                    'E_4-Situação da Unidade - Anterior',
                    'E_4-Estrutura da Construção - Anterior',
                    'E_4-Cobertura - Anterior',
                    'E_4-Paredes - Anterior',
                    'E_4-Revestimento Fachada - Anterior',
                    #'E_4-Valor Venal da Edificação - Anterior',
                    #'E_4-Valor Venal do Terreno - Anterior',
                    'E_4-Valor Venal do Imóvel - Anterior',
                    'E_4-Valor do Imposto - Anterior',

                    'E_4-Tipo do Imóvel - Atual',
                    'E_4-Área Terreno - Atual',
                    'E_4-Uso do Imóvel - Atual',
                    'E_4-Área Construída Unidade - Atual',
                    'E_4-Área Total Edificada - Atual',
                    'E_4-Fração Ideal - Atual',
                    'E_4-Situação da Quadra - Atual',
                    'E_4-Topografia - Atual',
                    'E_4-Pedologia - Atual',
                    'E_4-Tipo de Construção - Atual',
                    'E_4-Alinhamento - Atual',
                    'E_4-Situação da Edificação - Atual',
                    'E_4-Situação da Unidade - Atual',
                    'E_4-Estrutura da Construção - Atual',
                    'E_4-Cobertura - Atual',
                    'E_4-Paredes - Atual',
                    'E_4-Revestimento Fachada - Atual',
                    #'E_4-Valor Venal da Edificação - Atual',
                    #'E_4-Valor Venal do Terreno - Atual',
                    'E_4-Valor Venal do Imóvel - Atual',
                    'E_4-Valor do Imposto - Atual',
                    #'E_4-Diferença',

                    'E_5-Exercício',
                    'E_5-Tipo do Imóvel - Anterior',
                    'E_5-Área Terreno - Anterior',
                    'E_5-Uso do Imóvel - Anterior',
                    'E_5-Área Construída Unidade - Anterior',
                    'E_5-Área Total Edificada - Anterior',
                    'E_5-Fração Ideal - Anterior',
                    'E_5-Situação da Quadra - Anterior',
                    'E_5-Topografia - Anterior',
                    'E_5-Pedologia - Anterior',
                    'E_5-Tipo de Construção - Anterior',
                    'E_5-Alinhamento - Anterior',
                    'E_5-Situação da Edificação - Anterior',
                    'E_5-Situação da Unidade - Anterior',
                    'E_5-Estrutura da Construção - Anterior',
                    'E_5-Cobertura - Anterior',
                    'E_5-Paredes - Anterior',
                    'E_5-Revestimento Fachada - Anterior',
                    #'E_5-Valor Venal da Edificação - Anterior',
                    #'E_5-Valor Venal do Terreno - Anterior',
                    'E_5-Valor Venal do Imóvel - Anterior',
                    'E_5-Valor do Imposto - Anterior',

                    'E_5-Tipo do Imóvel - Atual',
                    'E_5-Área Terreno - Atual',
                    'E_5-Uso do Imóvel - Atual',
                    'E_5-Área Construída Unidade - Atual',
                    'E_5-Área Total Edificada - Atual',
                    'E_5-Fração Ideal - Atual',
                    'E_5-Situação da Quadra - Atual',
                    'E_5-Topografia - Atual',
                    'E_5-Pedologia - Atual',
                    'E_5-Tipo de Construção - Atual',
                    'E_5-Alinhamento - Atual',
                    'E_5-Situação da Edificação - Atual',
                    'E_5-Situação da Unidade - Atual',
                    'E_5-Estrutura da Construção - Atual',
                    'E_5-Cobertura - Atual',
                    'E_5-Paredes - Atual',
                    'E_5-Revestimento Fachada - Atual',
                    #'E_5-Valor Venal da Edificação - Atual',
                    #'E_5-Valor Venal do Terreno - Atual',
                    'E_5-Valor Venal do Imóvel - Atual',
                    'E_5-Valor do Imposto - Atual',
                    #'E_5-Diferença',

                    'E_6-Exercício',
                    'E_6-Tipo do Imóvel - Anterior',
                    'E_6-Área Terreno - Anterior',
                    'E_6-Uso do Imóvel - Anterior',
                    'E_6-Área Construída Unidade - Anterior',
                    'E_6-Área Total Edificada - Anterior',
                    'E_6-Fração Ideal - Anterior',
                    'E_6-Situação da Quadra - Anterior',
                    'E_6-Topografia - Anterior',
                    'E_6-Pedologia - Anterior',
                    'E_6-Tipo de Construção - Anterior',
                    'E_6-Alinhamento - Anterior',
                    'E_6-Situação da Edificação - Anterior',
                    'E_6-Situação da Unidade - Anterior',
                    'E_6-Estrutura da Construção - Anterior',
                    'E_6-Cobertura - Anterior',
                    'E_6-Paredes - Anterior',
                    'E_6-Revestimento Fachada - Anterior',
                    #'E_6-Valor Venal da Edificação - Anterior',
                    #'E_6-Valor Venal do Terreno - Anterior',
                    'E_6-Valor Venal do Imóvel - Anterior',
                    'E_6-Valor do Imposto - Anterior',

                    'E_6-Tipo do Imóvel - Atual',
                    'E_6-Área Terreno - Atual',
                    'E_6-Uso do Imóvel - Atual',
                    'E_6-Área Construída Unidade - Atual',
                    'E_6-Área Total Edificada - Atual',
                    'E_6-Fração Ideal - Atual',
                    'E_6-Situação da Quadra - Atual',
                    'E_6-Topografia - Atual',
                    'E_6-Pedologia - Atual',
                    'E_6-Tipo de Construção - Atual',
                    'E_6-Alinhamento - Atual',
                    'E_6-Situação da Edificação - Atual',
                    'E_6-Situação da Unidade - Atual',
                    'E_6-Estrutura da Construção - Atual',
                    'E_6-Cobertura - Atual',
                    'E_6-Paredes - Atual',
                    'E_6-Revestimento Fachada - Atual',
                    #'E_6-Valor Venal da Edificação - Atual',
                    #'E_6-Valor Venal do Terreno - Atual',
                    'E_6-Valor Venal do Imóvel - Atual',
                    'E_6-Valor do Imposto - Atual']
                    #'E_6-Diferença']
    current_year = datetime.now().year

    exercicio_lancamento = {
        current_year - 5: "E_1-",
        current_year - 4: "E_2-",
        current_year - 3: "E_3-",
        current_year - 2: "E_4-",
        current_year - 1: "E_5-",
        current_year: "E_6-",
    }

    grouped_data = {}
    for index, row in csv_data.iterrows():
        matricula = row['Matrícula']
        exercicio = row['Exercício']
        cabecalho = exercicio_lancamento[exercicio]
        if matricula not in grouped_data:
            data_dict = {}
            for key in header_output:
                data_dict[key] = None
            grouped_data[matricula] = {}

        # Cria um dicionário para armazenar os dados do arquivo
        for header in header_output:
            if header.startswith('E_'):
                short_header = header[4:]  # Descarta as primeiras 14 letras do header
                if header[0:4] == cabecalho:
                    data_dict[header] = row[short_header]
            else:
                data_dict[header] = row[header]
            data_dict.update(data_dict)

        grouped_data[matricula].update(data_dict)

        grouped_data_df = pd.DataFrame.from_dict(grouped_data, orient='index')
    return grouped_data_df


def gera_notificacao_impug(csv_data):
    #DEFINE O TIPO DE LAYOUT QUE SERÁ USADO (10000 = PARA LANÇAMENTO IGUAL E 100002 PARA LANÇAMENTO DIFERENTE)
    #10000 FOI O NOME DADO AO MODELO DE NOTIFICAÇÃO DE IMPUGNAÇÃO COM NOVO LANÇAMENTO IGUAL AO ANTERIOR
    for index, row in csv_data.iterrows():
        valor_anterior = row['Valor do Imposto - Anterior']
        valor_atual = row['Valor do Imposto - Atual']
    
        if valor_anterior == valor_atual:
            csv_data.at[index, 'Modelo do Layout'] = 10000
        else:
            csv_data.at[index, 'Modelo do Layout'] = 10002

    csv_data['Modelo do Layout'] = csv_data['Modelo do Layout'].astype(int)
    
    return csv_data


def verifica_auditor(grouped_data_df):
    # VERIFICA SE O USUÁRIOS QUE FEZ O LANÇAMENTO É AUDITOR, SE NÃO FOR COLOCA O NOME DO GERENTE DA GIPTU 
    grouped_data_df['Código do Objeto Cliente'] = range(1, len(grouped_data_df) + 1)
    chave, valor = GERENTE.popitem()
    auditor = chave
    auditor_matricula = valor
    for index, row in grouped_data_df.iterrows():
        if row['Usuário'] not in AUDITORES:
            grouped_data_df.at[index, 'Usuário'] = auditor
            grouped_data_df.at[index, 'Matrícula_Usuário'] = auditor_matricula

    return grouped_data_df


def gera_tabela(csv_data, matriz):
    # GERA A TABELA QUE SERÁ UTILIZADA PARA MALA DIRETA DO WORD + O ARQUIVO TXT QUE SERÁ ENTREGUE AOS CORREIOS
    # AQUI NO STM AO CLICAR EM EMITIR NOTIFICAÇÃO GERARIA A NOTIFICAÇÃO CORRETO (3 TIPOS DIFERENTES) E O ARQUIVO TXT PARA DOWNLOAD (FUTURAMENTE DEVERÁ CONVERSAR COM API)
    # DEVE HAVER UMA FORMA DE GERAR NOTIFICAÇÃO EM LOTE. GERAR VÁRIOS HTML SEPARADAMENTE PARA NÃO BAGUNÇAR A FORMATAÇÃO E JUNTÁ-LOS NÃO SEI COMO.
    # -> QUESTIONAR O USUÁRIO SE DESEJA EMITIR UM ÚNICO PDF OU SEPARADO. DOWNLOAD DE UM ARQUIVO .RAR OU .ZIP
    output_file_path = f"e-Carta_{matriz}_servico.txt"
    csv_data.to_csv(output_file_path, sep='|', encoding='utf-8', index=False, decimal=',', header=False)
    
    #LEGENDA: TIPO DE REGISTRO PREENCHIDO COM 1|NÚMERO DO LOTE DO OUTRO ARQUIVO O QUAL DEVE SER VERIFICADO|AUTORIZAÇÃO = A
    conteudo = f"1|{lote}|A"
    nome_arquivo = f"e-Carta_{matriz}_Resposta27012023073701.txt"
    with open(nome_arquivo, 'w') as arquivo:
        arquivo.write(conteudo)

    if matriz == '37642_1':
        filename = "tabela_impug.xlsx"
    else:
        filename = "tabela.xlsx"
        
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        csv_data.to_excel(writer, index=False)
    return


def verificar_numero_ou_vazio(valor):
    # IRRELEVANTE PARA STM
    if valor.strip() == '' or (valor.strip().isdigit() and int(valor.strip()) < 2):
        return False
    else:
        return True
    
def numero_lote(valor):
    # IRRELEVANTE PARA STM
    if valor.strip().isdigit():
        return False
    else:
        return True
    
def copiar_arquivos(tipo):
    # IRRELEVANTE PARA O STM
    notificacao = "not_lanc.docm"
    if tipo == '1':
        notificacao = "not_lanc-impug.docm"

    notificacao_caminho = r"Q:\\GIPTU\\AUDITORIA\\CLEBER\\08-AUTOMACAO\\02-notificacao\\" + notificacao
    #notificacao_caminho = "C:\\Users\\02166642179\\Desktop\\Python\\NotificacaoIPTU\\" + notificacao
    incl = r"Q:\\GIPTU\\AUDITORIA\\CLEBER\\08-AUTOMACAO\\05-tramitacao\\inc_not-tramit-siged.py"

    try:
        copyfile(incl, os.getcwd() + '\\' + '2_inc_not-tramit-siged.py')
        #copyfile(notificacao_caminho, os.getcwd() + '\\resultado\\' + notificacao)
        copyfile(notificacao_caminho, os.getcwd() + '\\' + notificacao)
    except Exception as e:
        print (e)

    return

####### NOVOS AUDITORES DEVEM SER INCLUÍDOS AQUI #########
AUDITORES = {'CLEBER TONELLO PEDRO JUNIOR': '1402609A',
             'CARLOS HENRIQUE MARTINS REZENDE': '1419706A',
             'THIAGO NORONHA DAMASCENO OLIVEIRA': '1425161A',
             'JOAO LUIZ MENDES ROMAO': '1403079A',
             'ARMANDO CLÁUDIO SIMÕES DA SILVA': '0629499A'
             }

############ DEVE CORRESPONDER À GERÊNCIA ATUAL ##############
GERENTE = {'JOAO LUIZ MENDES ROMAO': '1403079A'}

# Obtém o caminho absoluto do arquivo CSV na pasta atual
current_directory = os.getcwd()
# Considera apenas o primeiro arquivo CSV encontrado
csv_files = [file for file in os.listdir(current_directory) if file.endswith('.csv')]
csv_file_path = os.path.join(current_directory, csv_files[0])

# Verifica se há pelo menos um arquivo CSV na pasta atual
if len(csv_files) == 0:
    print("Nenhum arquivo CSV encontrado na pasta atual.")
    exit()


########################################## A PARTIR DAQUI ###############################################
################ COMENTAR ESSE TRECHO ASSIM QUE NÃO HOUVER MAIS NOTIFICAÇÃO DE IMPUGNAÇÃO ###############
tipo = input("Notificação de Impugnação? 1-Sim / 0-Não ou Aperte ENTER \n")
while verificar_numero_ou_vazio(tipo):
    tipo = input("Notificação de Impugnação? 1-Sim / 0-Não ou Aperte ENTER \n")
#########################################################################################################

lote = input("Informar o número do lote da notificação\n")
while numero_lote(lote):
    lote = input("Informar o número do lote da notificação\n")

with open(csv_file_path, 'r', encoding='utf-8') as csv_file:
    # Lê o arquivo CSV com o pandas
    csv_data = pd.read_csv(csv_file, delimiter='|', decimal=',', dtype={'Cartão de Postagem': str})
    csv_data.sort_values(by=['Matrícula', 'Exercício'], inplace=True)
    ###################################################################################
    # SE A DIFERENÇA ESTIVER VINDO CORRETAMENTE E O NOME DA COLUNA MATRÍCULA (USUÁRIO) TIVER SIDO ALTERADO ESSE TRECHO DEVE SER REMOVIDO
    csv_data.rename(columns={csv_data.columns[88]: 'Matrícula_Usuário'}, inplace=True)
    csv_data['Diferença'] = csv_data['Diferença'] * (-1)
    ###################################################################################
    if tipo == '1':
        grouped_data_df = gera_notificacao_impug(csv_data)
        matriz = '37642_1'
    else:
        grouped_data_df = gera_notificacao(csv_data)
        matriz = 'XXXXX_X' # A SER DEFINIDA JUNTO AOS CORREIOS.
    
    grouped_data_df['Número do Lote'] = csv_data['Número do Lote'].replace(1, lote)

        
verifica_auditor(grouped_data_df)

gera_tabela(grouped_data_df, matriz)

copiar_arquivos(tipo)