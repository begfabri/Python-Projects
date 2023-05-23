# Crea file XML di istruzione taglio automatico 


Il presente programma viene utilizzata in una azienda leader nel settore delle porte da garage. Nello specifico serve all'ottimizzazione del processo di taglio dei pannelli sandwich utilizzati nelle porte sezionali. Per l'ottimizzazione del taglio dei pannelli, viene utilizzato un software di terze parti che, per eseguire l'operazione, richiede in ingresso un file in formato XML. Tale file contiene la lista delle porte completa delle specifiche delle lavorazioni da eseguire sui pannelli che compongono la porta sezionale. 


# Descrizione casistica #

Nello specifico si tratta di estrarre una lista di codici di porte, presenti su gestionale Galileo (AS400), ed utilizzare questa lista di codici per estrarre i dati tecnici di configurazione delle porte da un database MS-SQL.

Come prima operazione, tramite il file importaLotti.py si occupa di reperire la lista dei codici dal gestionale Galileo, filtrando i dati in base al lotto di produzione. Vista la necessità di interfacciare due database differenti, lo stesso file si occuperà di inserisce i dati estratti in una nuova tabella creata nel database MS-SQL, che sarà utilizzata per restituire la lista degli articoli.

Una volta estratti i dati da Galileo, si provvede, tramite il file CreaXml.py, all'estrazione dei dati dal database MS-SQL ed a generare la struttura del file Catalog.XML, da inviare al software di terze parti. 

# Esempio di file XML generato #
![outputFile](https://user-images.githubusercontent.com/52453317/177986719-0d3c60b2-e316-4aa8-8b04-244782e82499.JPG)
