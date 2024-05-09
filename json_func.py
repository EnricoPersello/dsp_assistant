import json
import os


# Funzione per leggere file config
def read_config():
    # Legge il file con le credenziali
    # Ottieni la directory dello script Python
    DIRECTORY_SCRIPT = os.path.dirname(os.path.abspath(__file__))

    # Costruisci il percorso completo per il file config.json
    PERCORSO_CONFIG = os.path.join(DIRECTORY_SCRIPT, "config.json")

    # Leggi le credenziali dal file config o creane uno
    while True:
        # Apre il file config se presente
        try:
            with open(PERCORSO_CONFIG, "r") as CONFIG_FILE:
                CONFIG_DATA = json.load(CONFIG_FILE)

                EMAIL = CONFIG_DATA.get("email")
                PSW = CONFIG_DATA.get("password")
                CORTEX = CONFIG_DATA.get("cortex")
                DOWNLOAD = CONFIG_DATA.get("download_path")
                PERCORSO_STOP_ORARI = CONFIG_DATA.get("stop_orari_path")
                PERCORSO_STORICO_INGRESSI = CONFIG_DATA.get("storico_stop_orari_path")

                if (
                    not EMAIL
                    or not PSW
                    or not CORTEX
                    or not DOWNLOAD
                    or not PERCORSO_STOP_ORARI
                    or not PERCORSO_STORICO_INGRESSI
                ):
                    raise KeyError

                break

        # Se non ha trovato il file o mancano dati crea un nuovo file config
        except (FileNotFoundError, json.JSONDecodeError, KeyError):
            create_config()

    return EMAIL, PSW, CORTEX, DOWNLOAD, PERCORSO_STOP_ORARI, PERCORSO_STORICO_INGRESSI


# Funzione per creare file config
def create_config():
    # Chiede tutte le cose da salvare nel file
    print(
        "Non ho trovato il file di configurazione o dei dati non sono corretti, facciamone uno nuovo\n"
    )
    EMAIL = input("Inserisci l'email: ")
    PASSWORD = input("Inserisci la password: ")
    DOWNLOAD_PATH = input("Inserisci il percorso della cartella download: ")
    STOP_ORARI_PATH = input("Inserisci il persorso del file stop orari: ")
    STORICO_STOP_ORARI_PATH = input(
        'Inserisci il percorso della cartella "storico" dei file stop orari: '
    )

    # Controlla la formattazione delle stringhe
    DOWNLOAD_PATH = DOWNLOAD_PATH.replace('"', "")
    STOP_ORARI_PATH = STOP_ORARI_PATH.replace('"', "")
    STORICO_STOP_ORARI_PATH = STORICO_STOP_ORARI_PATH.replace('"', "")

    # Salva le informazioni in un file di configurazione JSON
    CONFIG_DATA = {
        "email": EMAIL,
        "password": PASSWORD,
        "cortex": "https://logistics.amazon.it/operations/execution/",
        "download_path": DOWNLOAD_PATH,
        "stop_orari_path": STOP_ORARI_PATH,
        "storico_stop_orari_path": STORICO_STOP_ORARI_PATH,
    }

    with open("config.json", "w") as CONFIG_FILE:
        json.dump(CONFIG_DATA, CONFIG_FILE)
