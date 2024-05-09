from time import sleep
from webbrowser import Chrome
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
)

import xpath


# Funzione per leggere del testo, ritorna il testo letto
def leggi_testo(driver, pausa, xpath):
    return (
        WebDriverWait(driver, pausa)
        .until(EC.visibility_of_element_located((By.XPATH, xpath)))
        .text
    )


# Funzione per cliccare, non ritorna niente
def click(driver, pausa, xpath):
    WebDriverWait(driver, pausa).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    ).click()


# Funzione per avviare cortex su chrome
def prepara_cortex(link, email, psw, pausa, day):
    # Scarca il driver di Chrome e apri la finestra
    chrome_driver = ChromeDriverManager().install()
    driver = Chrome(service=Service(chrome_driver))
    driver.maximize_window()

    # Vai al link
    driver.get(link)

    # Inserisce le credenziali con delle piccole attese statiche per cercare di ingannare il rilevatore di robot
    sleep(1)
    driver.find_element(By.ID, "ap_email").send_keys(email)
    sleep(5)
    driver.find_element(By.ID, "ap_password").send_keys(psw)
    sleep(2)
    driver.find_element(By.ID, "signInSubmit").click()
    sleep(1)

    # Fa il primo tentativo in automatico
    try:
        # Seleziona la ds
        driver.get(
            f"https://logistics.amazon.it/operations/execution/itineraries?operationView=true&selectedDay={day}&serviceAreaId=36574a56-d035-4cd6-a11f-31712acc4ef6"
        )
    except TimeoutException:
        # Prova a rifare login
        try:
            driver.back()
            WebDriverWait(driver, pausa).until(
                EC.presence_of_element_located((By.ID, "ap_password"))
            ).send_keys(psw)
            WebDriverWait(driver, pausa).until(
                EC.element_to_be_clickable((By.ID, "signInSubmit"))
            ).click()
        # Fail e tentativo manuale
        except:
            # print("\nQuesto è il link di cortex se ti serve: ", link)
            input(
                "Esegui il login manualmente e vai alla pagina principale con le rotte, poi premi invio"
            )

    return driver
