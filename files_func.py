import os


# Controlla esistenza cartelle per salvataggio file storico stop orari
def controlla_cartelle(path_storico_ingressi, anno, mese):
    percorso_anno = os.path.join(path_storico_ingressi, anno)
    percorso_mese = os.path.join(percorso_anno, mese)

    if not os.path.exists(percorso_anno):
        os.makedirs(percorso_anno)
    else:
        pass

    if not os.path.exists(percorso_mese):
        os.makedirs(percorso_mese)
    else:
        pass
