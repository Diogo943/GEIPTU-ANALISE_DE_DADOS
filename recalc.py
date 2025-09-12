import subprocess

required_libraries = ['selenium', 'webdriver_manager', 'xlsxwriter', 'shutil']

for library in required_libraries:
    try:
        __import__(library)
    except ImportError:
        subprocess.check_call(['pip', 'install', library])

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
from time import sleep
from operator import itemgetter
from itertools import groupby
from pathlib import Path
from shutil import copyfile
import os


def is_download_finished(temp_folder):
    firefox_temp_file = sorted(Path(temp_folder).glob('*.part'))
    chrome_temp_file = sorted(Path(temp_folder).glob('*.crdownload'))
    downloaded_files = sorted(Path(temp_folder).glob('*.*'))
    if (len(firefox_temp_file) == 0) and \
       (len(chrome_temp_file) == 0) and \
       (len(downloaded_files) >= 1):
        return True
    else:
        return False
    

def conf_impug():
    loop = False
    while not loop:
        try:
            impug = (int(input('É revisão de impugnação? 1-Sim / 2-Não ')))
            loop = True
        except ValueError:
            print('Valor inserido não é 0 ou 1.')
    return impug


def conf_lote():
    loop = False
    while not loop:
        try:
            lote = (int(input('Deseja usar recálculo em lote? 1-Sim / 2-Não ')))
            loop = True
        except ValueError:
            print('Valor inserido não é 0 ou 1.')
    return lote


def recalc_lote(matriculas):
    ranges = []
    for k,g in groupby(enumerate(matriculas),lambda x:x[0]-x[1]):
        group = (map(itemgetter(1),g))
        group = list(map(int,group))
        ranges.append((group[0],group[-1]))

    for group in (ranges):
        driver.get("https://stm.manaus.am.gov.br/stm/servlet/hwmtmprocloterecalculoiptu")
        sleep(2)
        btn_insert = driver.find_element(By.NAME, "GX_BTNINSERT")
        btn_insert.click()

        #escrevendo as matrículas
        matriculas_ini = wait.until(EC.presence_of_element_located((By.NAME, "vINICIOPRICTIID")))
        matriculas_end = wait.until(EC.presence_of_element_located((By.NAME, "vFINALPRICTIID")))
        matriculas_ini.send_keys(group[0])
        matriculas_end.send_keys(group[1])
        sleep(2)

        #anos
        select_ini = Select(driver.find_element(By.NAME, "vINICIOTRBID"))
        select_end = Select(driver.find_element(By.NAME, "vFINALTRBID"))
        year_start = str(YEAR[0])
        year_end = str(YEAR[-1])
        select_ini.select_by_visible_text(year_start)
        select_end.select_by_visible_text(year_end)
        sleep(2)

        #processo
        process_number_form = wait.until(EC.presence_of_element_located((By.NAME, "vPRINMRPROCESSO")))
        process_number_form.send_keys(PROCESS_NUMBER)
        sleep(1)
        year_form = driver.find_element(By.NAME, "vPRIPROCESSOANO")
        year_form.send_keys(PROCESS_NUMBER[0:4])
        sleep(1)

        enter_button = wait.until(EC.element_to_be_clickable((By.NAME, "BTN_ENTER")))
        enter_button.click()
        sleep(5)

        wait.until_not(EC.visibility_of_element_located((By.CLASS_NAME, "gx-mask")))
    
    return

def recalc(matriculas):
    for id in matriculas:
        for year in YEAR:
            driver.get("https://stm.manaus.am.gov.br/stm/servlet/hwtprocrecalculoiptu?INS,0")
            
            id_form_insert = wait.until(EC.presence_of_element_located((By.NAME, "CTLPRICTIID1")))
            sleep(2)

            id_form_insert.send_keys(id)

            #impugnação
            if impug == 1:
                id_impug = Select(driver.find_element(By.NAME, "CTLPRIEHIMPUGNACAO"))
                id_impug.select_by_visible_text("SIM")

            process_number_form = wait.until(EC.presence_of_element_located((By.NAME, "CTLPRINMRPROCESSO")))
            sleep(2)

            process_number_form.send_keys(PROCESS_NUMBER)
            sleep(1)
            year_form = driver.find_element(By.NAME, "CTLPRIPROCESSOANO")
            year_form.send_keys(PROCESS_NUMBER[0:4])
            sleep(1)

            select = Select(driver.find_element(By.NAME, "vTRBID"))
            select.select_by_visible_text(str(year))

            enter_button = wait.until(EC.element_to_be_clickable((By.NAME, "BTN_ENTER")))
            enter_button.click()

            wait.until(EC.text_to_be_present_in_element((By.ID, "span_vCTIID"), str(id)))
            sleep(4)
            enter_button2 = driver.find_element(By.CLASS_NAME, "BtnConfirmar")
            enter_button2.click()

            ### SIMULAÇÃO ###
            simulation_insert_button = wait.until(EC.element_to_be_clickable((By.NAME, "GX_BTNINSERT")))
            simulation_insert_button.click()
            sleep(1)
            wait.until_not(EC.visibility_of_element_located((By.CLASS_NAME, "gx-mask")))
                    
            while driver.find_element(By.ID, "span_vTOTALCALCULADO").text == "0,00":
                simulation_insert_button.click()
                sleep(10)

            enter_button3 = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "BtnConfirmar")))
            enter_button3.click()
            sleep(3)
            wait.until_not(EC.visibility_of_element_located((By.CLASS_NAME, "gx-mask")))

########## LOGIN ###########
USER = "LOGIN"
PASSWORD = "SENHA"
############################

data = {}
filename = "dados_processo.txt"
impug = conf_impug()
lote = 0
if impug != 1:
    lote = conf_lote()

itens_interesse = ['ano', 'processo', 'matricula']
with open(filename, 'r', encoding='utf-8-sig') as f:
    for line in f:
        name, attr = line.strip().split("=")
        if name in itens_interesse:
            data[name] = attr

YEAR = data["ano"].replace(' ', '').split(',')
ids = data["matricula"].replace(' ', '').split(',')
matriculas = [int(x) for x in ids]
matriculas.sort()
#OBS = data["obs"]
PROCESS_NUMBER = data["processo"]

service = FirefoxService(executable_path=GeckoDriverManager().install())
print("Driver Installed")

#preferences
current_dir = os.getcwd()
download_dir = current_dir

# Set the Firefox profile to automatically download files to the download directory
options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", download_dir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")

driver = webdriver.Firefox(service=service, options=options)

driver.get("https://stm.manaus.am.gov.br/stm/servlet/hwlogin")

#### LOGIN ####
wait = WebDriverWait(driver, 5000)

login_button = wait.until(EC.element_to_be_clickable((By.NAME, "BUTTON1")))
user_form = wait.until(EC.presence_of_element_located((By.NAME, "vUSULOGIN")))
password_form = wait.until(EC.presence_of_element_located((By.NAME, "vUSUSENHA")))

user_form.send_keys(USER)
password_form.send_keys(PASSWORD)
login_button.click()

if lote == 1:
    recalc_lote(matriculas)
else:
    recalc(matriculas)

    driver.get("https://stm.manaus.am.gov.br/stm/servlet/hwmtmprocrecalculoiptu")

    #listando os recálculos
    recal_list = wait.until(EC.presence_of_element_located((By.NAME, "vNMRPROCINI")))
    recal_list.send_keys(PROCESS_NUMBER)
    sleep(1)

    search_button = driver.find_element(By.NAME, "GX_BTNSEARCH")
    search_button.click() 
    sleep(1) 

    wait.until_not(EC.visibility_of_element_located((By.CLASS_NAME, "gx-mask")))

    #gerar csv  
    csv_button = driver.find_element(By.NAME, "BTNGERARCSV")
    csv_button.click()    

    #download
    sleep(3) 

    while not is_download_finished(download_dir):
        sleep(1)

driver.close()

print("Matrícula(s) recalculada(s)")

gera_notificacao_script = "gera_notificacao.py"
script_caminho = "Q:\\GIPTU\\AUDITORIA\\CLEBER\\08-AUTOMACAO\\02-notificacao\\" + gera_notificacao_script
copyfile(script_caminho, os.getcwd() + '\\' + gera_notificacao_script)

print("Arquivos de notificação copiados para a pasta atual")