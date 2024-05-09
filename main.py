import os
import locale
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from files_func import controlla_cartelle
from json_func import read_config
from web_func import prepara_cortex, leggi_testo

from case_a import case_a
import xpath


##############################################################################################################################################################
##############################################################################################################################################################
#########################################################################   SET UP   #########################################################################
##############################################################################################################################################################
##############################################################################################################################################################


print("SET UP DEL PROGRAMMA IN CORSO......\n")

# Leggi dal file config.json le credenziali e i link
EMAIL, PSW, CORTEX, DOWNLOAD, PERCORSO_STOP_ORARI, PERCORSO_STORICO_INGRESSI = (
    read_config()
)

# Costanti
PAUSA_LUNGA = 20
PAUSA = 10
PAUSA_CORTA = 2
FERMATE_TOT = []
DS = "DFV2"
POSIZIONE_DS = "37R2+76 Udine, Ente di decentramento regionale di Udine"
MAPS = "https://www.google.com/maps/"

# Variabili
richiesta = None
stop = None
file_ing_fatto = False
update_fermate = []
conta_pausa = []
first_stop_times = []
last_stop_times = []
it_problems = []
motivo_rip = []

# Costanti orarie e temporali
locale.setlocale(locale.LC_TIME, "it_IT.UTF-8")
ora_attuale = datetime.now().time()
ORA_DI_RIFERIMENTO = datetime.strptime("11:00:00", "%H:%M:%S").time()
ANNO = datetime.now().today().strftime("%Y")
MESE_CORRENTE = datetime.now().strftime("%B").upper()
OGGI = datetime.now().today().strftime("%d.%m.%Y")
OGGI_INGRESSI = datetime.today().strftime("%d-%m")
DAY_CORTEX = datetime.today().strftime("%Y-%m-%d")


# Percorsi composti
INGRESSI_CUBAGGI = f"Ingressi e cubaggi DBLS {OGGI_INGRESSI}.xlsx"
INGRESSI_CUBAGGI_PLUS = INGRESSI_CUBAGGI.replace(" ", "+")
PERCORSO_INGRESSI_AMZ = os.path.join(DOWNLOAD, INGRESSI_CUBAGGI)
PERCORSO_INGRESSI_AMZ_PLUS = os.path.join(DOWNLOAD, INGRESSI_CUBAGGI_PLUS)

NOME_FILE_STOP_ORARI_OGGI = f"CONTEGGIO STOP ORARI {OGGI}.xlsx"
PERCORSO_STOP_ORARI_OGGI = os.path.join(
    PERCORSO_STORICO_INGRESSI, ANNO, MESE_CORRENTE, NOME_FILE_STOP_ORARI_OGGI
)


# Controlla che ci siano le cartelle per salvare i file sul pc
controlla_cartelle(
    path_storico_ingressi=PERCORSO_STORICO_INGRESSI, anno=ANNO, mese=MESE_CORRENTE
)

# Apertura Chrome
TOT_DA = 0
DRIVER = prepara_cortex(link=CORTEX, email=EMAIL, psw=PSW, pausa=PAUSA, day=DAY_CORTEX)

##############################################################################################################################################################
##############################################################################################################################################################
###########################################################################   MAIN   #########################################################################
##############################################################################################################################################################
##############################################################################################################################################################


while True:
    print("\n\nCOSA VUOI FARE?")
    # print('scrivi "i" per info sul programma')
    print('scrivi "a" per compilare un nuovo file ingressi')
    # print('scrivi "b" per l\'analisi avanzata dei pacchi problematici')
    # print('scrivi "c" per avviare l\'aggiornamento automatico della tabella')
    # print('scrivi "d" per lo scarico dati dei resi')
    print('scrivi "q" per chiudere il programma')
    richiesta = input("Poi premi invio!\nLa tua scelta: ").lower()
    print("\n")

    # Se il numero dei DA è 0 trova il numero totale delle rotte
    # Non è presente l'else perchè il ciclo si ripete finchè non trova un numero >0
    if TOT_DA == 0:
        try:
            TOT_DA = int(
                leggi_testo(driver=DRIVER, pausa=PAUSA, xpath=xpath.XPATH_TOT_DA)
            )
        except TimeoutException:
            print("NON HO TROVATO NESSUNA ROTTA")
    # print(TOT_DA)

    # Se ha già controllato il numero delle rotte esegui la richiesta
    if TOT_DA > 0:
        match richiesta:

            # Compilazione nuovo file ingressi o apertura esistente
            case "a":
                FERMATE_TOT = case_a(
                    percorso_stop_orari_oggi=PERCORSO_STOP_ORARI_OGGI,
                    pausa_corta=PAUSA_CORTA,
                    pausa=PAUSA,
                    fermate_tot=FERMATE_TOT,
                    maps=MAPS,
                    posizione_ds=POSIZIONE_DS,
                    tot_da=TOT_DA,
                    percorso_ingressi_amz=PERCORSO_INGRESSI_AMZ,
                    percorso_ingressi_amz_plus=PERCORSO_INGRESSI_AMZ_PLUS,
                    percorso_stop_orari=PERCORSO_STOP_ORARI,
                )

            case "q":
                break


DRIVER.quit()
