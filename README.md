# ExcelEmailer

**ExcelEmailer** è uno strumento semplice e intuitivo per inviare email di massa utilizzando un file Excel come sorgente. Supporta allegati PDF opzionali e configurazione rapida tramite console.

## Funzionalità
- Lettura di indirizzi email da un file Excel.
- Invio di email personalizzate a più destinatari.
- Possibilità di allegare un file PDF a ogni email.
- Compatibile con account Gmail (e altri provider con configurazioni SMTP simili).

##Utilizzo
-Avvia il programma.
-Fornisci i seguenti input:
-Percorso del file Excel.
-Oggetto e corpo dell'email.
-(Opzionale) Percorso del file PDF da allegare.
-Credenziali email del mittente (es. Gmail).
-Il programma invierà un'email a tutti gli indirizzi trovati nel file Excel e mostrerà i risultati in console.


Il programma utilizza Gmail con porta 587 e SSL abilitato. Per altri provider, potrebbe essere necessario modificare le impostazioni SMTP.
Gmail richiede l'abilitazione dell'accesso alle app meno sicure o la creazione di una password specifica per l'app.
Contribuzione
Le richieste di miglioramento e le segnalazioni di bug sono benvenute! Sentiti libero di creare un issue o una pull request.
