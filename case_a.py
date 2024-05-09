from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from web_func import click, leggi_testo
import xlwings as xw
import pandas as pd
import xpath


def case_a(
    percorso_stop_orari_oggi,
    driver,
    pausa_corta,
    pausa,
    fermate_tot,
    maps,
    posizione_ds,
    tot_da,
    percorso_ingressi_amz,
    percorso_ingressi_amz_plus,
    percorso_stop_orari,
):
    # Controlla se il file è già stato creato
    try:
        xw.Book(percorso_stop_orari_oggi)

    # Se non lo trova crea uno nuovo
    except FileNotFoundError:
        print(
            "Non ho trovato nessun file salvato con la data di oggi quindi ne creo uno nuovo"
        )
        print("\nScarico i dati relativi alle fermate")

        # Ottiene da cortex la lista col numero di stop di ogni rotta
        if len(fermate_tot) == 0:
            fermate_tot = lista_fermate(driver=driver, pausa_corta=pausa_corta)

        print("Calcolo i transit time")

        # Crea le liste col transit time per il primo e ultimo stop
        if len(first_stop_times) == 0:
            first_stop_times, last_stop_times = transit_time(
                driver=driver,
                pausa=pausa,
                pausa_corta=pausa_corta,
                fermate=fermate_tot,
                link_maps=maps,
                posizione_sede=posizione_ds,
                totale_driver=tot_da,
            )

        print(
            "Scrivo i dati ottenuti su un nuovo file Excel e lo salvo con la data di oggi"
        )

        # Apre il file ingressi fornito da amz e crea il nostro file ingressi
        if file_ing_fatto == False:
            file_ing_fatto = scrivi_nuovo_file_ingressi(
                lista_fermate=fermate_tot,
                percorso_ingressi_amz=percorso_ingressi_amz,
                percorso_ingressi_amz_plus=percorso_ingressi_amz_plus,
                percorso_stop_orari=percorso_stop_orari,
                primi_stop=first_stop_times,
                ultimi_stop=last_stop_times,
                percorso_salvataggio=percorso_stop_orari_oggi,
            )

    return fermate_tot


##############################################################################################################################################################
##############################################################################################################################################################
########################################################################### STEP 1 ###########################################################################
##############################################################################################################################################################
##############################################################################################################################################################


# Crea la lista delle fermate totali di ogni rotta
def lista_fermate(driver, pausa_corta):
    # driver.refresh()
    elenco_fermate = []

    # Salva tutti gli elementi corrispondenti all'xpath indicato in una lista chiamata containers
    containers = WebDriverWait(driver, pausa_corta).until(
        EC.presence_of_all_elements_located((By.XPATH, xpath.XPATH_FERMATE))
    )

    # Per ogni elemento nella lista conteiners verifica se ha la parola fermate all'interno e
    # se sì salva il numero in una lista con l'ooportuna formattazione
    for container in containers:
        if "fermate" in container.text:
            fermate = container.text.replace("/", " ").split()
            elenco_fermate.append(int(fermate[1]))

    return elenco_fermate


##############################################################################################################################################################
##############################################################################################################################################################
########################################################################### STEP 2 ###########################################################################
##############################################################################################################################################################
##############################################################################################################################################################


# Cerca i transit time per ogni rotta e controllo pacchi "Mappa Imprecisa"
def transit_time(
    driver, pausa, pausa_corta, fermate, link_maps, posizione_sede, totale_driver
):

    # Funzione per inizializzare Maps
    def avvia_ricerca():

        # Definisci gli input e output
        partenza = WebDriverWait(driver, pausa).until(
            EC.presence_of_element_located((By.XPATH, xpath.XPATH_PARTENZA))
        )
        partenza.send_keys(posizione_sede)

        arrivo = WebDriverWait(driver, pausa).until(
            EC.presence_of_element_located((By.XPATH, xpath.XPATH_ARRIVO))
        )

        # Inverti è la freccia per invertire partenza e arrivo
        inverti = WebDriverWait(driver, pausa).until(
            EC.element_to_be_clickable((By.XPATH, xpath.XPATH_INVERTI))
        )

        return partenza, arrivo, inverti

    # Funzione per ottenere le stringhe col primo e ultimo indirizzo
    def get_primo_e_ultimo():

        # Clicca sulla cesella del driver
        WebDriverWait(driver, pausa).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    f'//*[@id="main"]/div/div/div[3]/div[1]/div[2]/div[{k}]/a/div/div/div',
                )
            )
        ).click()

        # Controlla se ci sono pacchi con mappa imprecisa
        try:
            it_mappa_imprecisa.append(
                mappa_imprecisa(xpath_it_problematici=xpath.XPATH_IT_PROBLEM_MAT)
            )

        except Exception:
            it_mappa_imprecisa.append(
                mappa_imprecisa(xpath_it_problematici=xpath.XPATH_IT_PROBLEM_POM)
            )

        # Cllick sulla casella del terzo stop
        try:
            try:
                click(driver=driver, pausa=pausa, xpath=xpath.XPATH_STOP_3_MAT)

            except ElementClickInterceptedException:
                driver.refresh()
                click(driver=driver, pausa=pausa, xpath=xpath.XPATH_STOP_3_MAT)

        except:
            click(driver=driver, pausa=pausa, xpath=xpath.XPATH_STOP_3_POM)

        # Registra il testo della fermata iniziale
        try:
            testo_1 = leggi_testo(
                driver=driver, pausa=pausa, xpath=xpath.XPATH_TESTO_FERMATA_MAT
            )

        except TimeoutException:
            testo_1 = leggi_testo(
                driver=driver, pausa=pausa, xpath=xpath.XPATH_TESTO_FERMATA_POM
            )

        # Torna alla lista delle fermate
        back_arrow = WebDriverWait(driver, pausa).until(
            EC.element_to_be_clickable((By.XPATH, xpath.XPATH_INDIETRO))
        )
        back_arrow.click()

        # Clicca sulla casella stop finale -1
        WebDriverWait(driver, pausa).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    f'//*[@id="main"]/div/div/div[3]/div[1]/div[3]/a[{fermate[i]-1}]/div/div/div',
                )
            )
        ).click()

        # Registra il testo della fermata finale
        try:
            testo_2 = leggi_testo(
                driver=driver, pausa=pausa, xpath=xpath.XPATH_TESTO_FERMATA_MAT
            )
        except TimeoutException:
            testo_2 = leggi_testo(
                driver=driver, pausa=pausa, xpath=xpath.XPATH_TESTO_FERMATA_POM
            )

        # Ritorna alla pagina principale
        back_arrow.click()
        back_arrow.click()

        # Restituisce gli indirizzi di partenza e arrivo
        return testo_1, testo_2

    # Funzione per trovare i pacchi con mappa imprecisa
    def mappa_imprecisa(xpath_it_problematici):
        it_problems = None

        try:
            # Controlla se c'è qualche stop con la dicitura mappa imprecisa
            elemento = WebDriverWait(driver, pausa_corta).until(
                EC.presence_of_element_located((By.XPATH, xpath.XPATH_MAPPA_IMPRECISA))
            )

            # Se è presente clicca sullo stop e registra l'IT
            if elemento.is_displayed():
                elemento.click()
                it_problems = leggi_testo(
                    driver=driver, pausa=pausa, xpath=xpath_it_problematici
                )
                driver.back()
                # Calcella l'elemento per evitare problemi con interazioni sucessive
                del elemento

        # Se non è presente vai oltre
        except TimeoutException:
            # print("qualcosa non ha funzionato")
            pass

        # Se sono stati trovati IT problemtici restituisci i valori
        if it_problems:
            return it_problems

    # Funzione per ricavere il trinsit time
    def get_transit(input_click, location):
        # Cerca la prima fermata
        input_click.send_keys(location)
        input_click.send_keys("\n")

        # Tempo prima fermata
        try:
            tempo_nec = leggi_testo(
                driver=driver, pausa=pausa_corta, xpath=xpath.XPATH_TRANSIT_TIME
            )

        # Riprova senza cambiare niente
        except TimeoutException:
            try:
                tempo_nec = leggi_testo(
                    driver=driver, pausa=pausa_corta, xpath=xpath.XPATH_TRANSIT_TIME
                )

            # Se non trova l'indirizzo prova a cercare e sfruttare i suggerimenti
            except TimeoutException:
                try:
                    input_click.clear()
                    input_click.clear()
                    input_click.send_keys(location)

                    click(
                        driver=driver, pausa=pausa_corta, xpath=xpath.XPATH_SUGG_RICERCA
                    )
                    tempo_nec = leggi_testo(
                        driver=driver, pausa=pausa_corta, xpath=xpath.XPATH_TRANSIT_TIME
                    )

                # Se ancora non trova il valore imposta zero
                except TimeoutException:
                    tempo_nec = "0 min"

        # Divide la stringa per salvare solo il numero
        tempo_nec = tempo_nec.split()
        # print(tempo_nec[0])
        tempo = tempo_nec[0]

        # Cancella tutto per evitare conflitti
        del tempo_nec
        input_click.clear()
        input_click.clear()
        inverti.click()

        # Restituisce il transit time trovato
        return tempo

    it_mappa_imprecisa = []
    it_mappa_imprecisa_corretta = []
    primo_stop = []
    ultimo_stop = []

    # Apri Maps
    driver.execute_script(f"window.open('{link_maps}', '_blank');")
    driver.switch_to.window(driver.window_handles[1])

    # Accetta coockies
    coockies = driver.find_elements(By.XPATH, '//*[@class="VfPpkd-vQzf8d"]')

    for coockie in coockies:
        if coockie.text == "Rifiuta tutto":
            coockie.click()
            break

    # Schiaccia sul pulsante per avviare la ricerca
    click(driver=driver, pausa=pausa, xpath=xpath.XPATH_AVVIA_RICERCA)

    # Definisci gli oggetti partenza, arrivo e inverti
    try:
        partenza, arrivo, inverti = avvia_ricerca()
    except TimeoutException:
        driver.refresh()
        partenza, arrivo, inverti = avvia_ricerca()

    # Ciclo per cercare tutti i transit time
    for i in range(totale_driver):
        # Vai a cortex
        driver.switch_to.window(driver.window_handles[0])

        # k serve per incrementare l'xpath della cesella del driver
        k = i + 2

        # Scarica le stringhe del primo e ultimo indirizzo
        inizio, fine = get_primo_e_ultimo()

        # Vai a maps
        driver.switch_to.window(driver.window_handles[1])

        # Trova i transit time di partenza e arrivo e aggiungili alla lista
        primo_stop.append(get_transit(input_click=arrivo, location=inizio))
        ultimo_stop.append(get_transit(input_click=partenza, location=fine))

    # driver.close()
    driver.switch_to.window(driver.window_handles[0])

    # Correggi la lista con gli IT con mappa imprecisa
    for i in range(len(it_mappa_imprecisa)):
        if it_mappa_imprecisa[i]:
            it_mappa_imprecisa_corretta.append(it_mappa_imprecisa[i])

    # Stampa gli IT dei pacchi con mappa imprecisa
    if len(it_mappa_imprecisa_corretta) != 0:
        print(
            "I PACCHI PROBLEMATICI TROVATI SONO:\n" + str(it_mappa_imprecisa_corretta)
        )

    # Restituisce le due liste con i transit time
    return primo_stop, ultimo_stop


##############################################################################################################################################################
##############################################################################################################################################################
########################################################################### STEP 3 ###########################################################################
##############################################################################################################################################################
##############################################################################################################################################################


# Scrive il file ingressi
def scrivi_nuovo_file_ingressi(
    lista_fermate,
    percorso_ingressi_amz,
    percorso_ingressi_amz_plus,
    percorso_stop_orari,
    primi_stop,
    ultimi_stop,
    percorso_salvataggio,
):
    # Definisci l'intervallo di celle da copiare
    def struttura(intervallo_finale, a, b):
        cella_partenza = a + str(intervallo_finale[0])
        cella_arrivo = b + str(intervallo_finale[len(intervallo_finale) - 1])
        tabella = cella_partenza + ":" + cella_arrivo
        # print(tabella)

        return tabella

    # Apre i due file e seleziona i relativi fogli di lavoro
    try:
        stop_orari_nuovo = xw.Book(percorso_stop_orari)
    except FileNotFoundError:
        print("Non ho trovato la versione da compilare del file conteggio stop orari")

    trovato = False
    while not trovato:
        try:
            ingressi_amz = xw.Book(percorso_ingressi_amz)
            trovato = True
        except FileNotFoundError:
            try:
                ingressi_amz = xw.Book(percorso_ingressi_amz_plus)
                trovato = True
            except FileNotFoundError:
                input(
                    "Non ho trovato gli ingressi mandati da Amazon, scarica il file poi premi invio"
                )

    ingressi = stop_orari_nuovo.sheets["INGRESSI"]
    dati = stop_orari_nuovo.sheets["DATI"]

    try:
        ws_part = ingressi_amz.sheets["Foglio1"]
    except:
        ws_part = ingressi_amz.sheets["Sheet1"]

    # Scrive il numero di fermate per ogni da
    for i in range(5, 5 + len(lista_fermate)):
        ingressi.cells(i, 7).value = int(lista_fermate[i - 5])

    # Scrive i transit time sul foglio dati
    for i in range(0, len(primi_stop)):
        dati.cells(i + 2, 3).value = int(primi_stop[i])
        dati.cells(i + 2, 4).value = int(ultimi_stop[i])

    # Nasconde la prima riga con le intestazioni
    ws_part.range(f"1:1").api.EntireRow.Hidden = True

    # Riconosce gli intervalli di celle non vuote
    first_visible_cell = ws_part.api.Cells.SpecialCells(
        12
    )  # 12 rappresenta xlCellTypeVisible
    intervalli = first_visible_cell.Address
    # print(intervalli)

    # Separa gli intervalli
    visible_cells = intervalli.split(",")
    # print(visible_cells)

    intervallo_finale = []
    for i in range(0, len(visible_cells)):
        stringa = visible_cells[i]
        stringa = stringa.replace("$", "")
        stringa = stringa.split(":")
        intervallo_finale.append(stringa[0])
        intervallo_finale.append(stringa[1])
        del stringa
    # print(intervallo_finale)

    # Nasconde le colonne che non servono
    ws_part.range(f"E:J").api.EntireColumn.Hidden = True

    # Elimina l'intervallo di celle alla fine della tabella
    lunghezza = len(intervallo_finale)
    del intervallo_finale[lunghezza - 1]
    del intervallo_finale[lunghezza - 2]
    # print(intervallo_finale)

    # Copia la tabella ottenuta con tutti i DA
    intervallo = struttura(intervallo_finale, "A", "D")
    ws_part.range(intervallo).copy()
    ws_part.range("P300").paste()

    # Numero di righe/DA
    diff = int(intervallo_finale[len(intervallo_finale) - 1]) - int(
        intervallo_finale[0]
    )

    # Determina un nuovo intervallo finale libero in modo da poter maneggiare i dati
    del intervallo_finale
    intervallo_finale = [300, 300 + diff]
    intervallo = struttura(intervallo_finale, "P", "S")
    tab = ws_part.range(intervallo).value

    # Ordina i valori in ordine di rotta
    df = pd.DataFrame(tab)
    df_ordinato = df.sort_values(
        by=0
    )  # Usa 0 per indicare la prima colonna senza intestazioni

    # Incolla la tabella sul file ingressi
    ingressi.range("A5").value = df_ordinato.values

    # Copia un valore a caso per svuotare gli appunti e poter ciudere il file ingerssi di amz
    ws_part.range("A1").copy()
    ingressi_amz.close()
    stop_orari_nuovo.save(percorso_salvataggio)

    return True
