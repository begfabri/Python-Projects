# Crea file XML di istruzione taglio automatico 


Il presente programma permette la creazione di un file XML, utilizzato per l'ottimizzazione del taglio pannelli sandwich eseguito da software di terze parti. Tale software accetta in ingresso un file XML, contentente le specifiche dei pannelli che compongono la porta pannellata. 


# Descrizione casistica #

Nello specifico si tratta di estrarre una lista di codici di porte da garage pannellate, presenti su gestionale Galileo (AS400), ed utilizzare questa lista di codici per estrarre i dati tecnici di configurazione delle porte da un database MS-SQL.

Come prima il programma, tramite il file importaLotti.py si occupa di reperire la lista dei codici dal gesionale Galileo filtrando i dati in base al lotto di produzione. Vista la necessità
di interfacciare due database differenti, lo stesso file inserisce i dati estratti in una nuova tabella creata nel database MS-SQL. Questo per permettere le sucessive elaborazioni dei dati.

Una volta estratti i dati da Galileo si provvede, tramite il file CreaXml.py, all'estrazione dei dati dal database MS-SQL ed a generare la struttura del file Catalog.XML, da inviare alla macchina di taglio dei pannelli.
