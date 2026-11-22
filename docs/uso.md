# Guida all'uso

Il convertitore legge il file delle letture esportato dal sistema Siemens (il file FC_report) e crea un file Excel con la ripartizione delle spese di riscaldamento e acqua tra gli appartamenti.

## Cosa ti serve

- **Il programma.** È un file `.exe` con un nome simile a `siemens-converter-windows-v0.2.0.exe`. Non va installato.
- **Il file FC_report.** È il file `.xls` esportato dal sistema Siemens. Il suo nome comincia con `FC_report`.
- **Il file dei dati statici** (facoltativo). È un file Excel `.xlsx` che compili tu con inquilini, millesimi, costi e letture precedenti. Come prepararlo è spiegato in [Il file dei dati statici](#il-file-dei-dati-statici).

## Scaricare il programma

1. Apri la [pagina dell'ultima versione](https://github.com/gjed/siemens-converter/releases/latest).
1. Nell'elenco `Assets` clicca sul file che termina con `.exe`.
1. Salva il file in una cartella che ritrovi facilmente, per esempio il Desktop.

## Convertire trascinando i file

### Solo il file FC_report

1. Apri la cartella che contiene il file FC_report.
1. Trascina il file FC_report sopra l'icona del programma.
1. Rilascia il file.
1. Aspetta la finestra con la scritta `Fatto!` e il nome del file creato.
1. Premi `OK`.

Senza il file dei dati statici il risultato è incompleto. Leggi [Se non usi il file dei dati statici](#se-non-usi-il-file-dei-dati-statici).

### Il file FC_report insieme al file dei dati statici

1. Metti il file FC_report e il file dei dati statici nella stessa cartella.
1. Clicca una volta sul file FC_report.
1. Tieni premuto il tasto `Ctrl` e clicca una volta sul file dei dati statici. Ora i due file sono selezionati.
1. Trascina i due file insieme sopra l'icona del programma.
1. Rilascia i file.
1. Aspetta la finestra con la scritta `Fatto!` e il nome del file creato.
1. Premi `OK`.

Non importa quale file selezioni per primo. Il programma li riconosce dall'estensione: il file `.xlsx` è quello dei dati statici, l'altro è il file FC_report.

Trascina al massimo due file. Il programma ignora quelli in più.

## Convertire con il doppio clic

1. Fai doppio clic sull'icona del programma.
1. Si apre la finestra `Seleziona FC_report (.xls)`. Clicca sul file FC_report.
1. Premi `Apri`.
1. Si apre la finestra `Seleziona dati statici (.xlsx) — o Annulla per saltare`. Clicca sul file dei dati statici. Se non hai il file dei dati statici, premi `Annulla` e salta il prossimo punto.
1. Premi `Apri`.
1. Aspetta la finestra con la scritta `Fatto!` e il nome del file creato.
1. Premi `OK`.

Se premi `Annulla` nella prima finestra, il programma si chiude senza creare nulla.

## Dove trovi il risultato

Il file creato si trova nella stessa cartella del file FC_report.

Il suo nome è `Riparto_<numero di serie>_<data>.xlsx`, per esempio `Riparto_TE99999999_15-01-2026.xlsx`. Il numero di serie e la data (giorno-mese-anno) sono quelli scritti dentro il file FC_report.

Se nella cartella c'è già un file con lo stesso nome, il programma lo sostituisce senza chiedere conferma. Se avevi modificato a mano quel file, prima salvane una copia con un altro nome.

## Cosa contiene il file creato

Il file ha cinque fogli.

| Foglio                          | Contenuto                                                                                                              |
| ------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| `Ripartizione_Termo 02.03.2026` | Il calcolo della ripartizione: letture iniziali e finali, consumi e costi di riscaldamento, acqua calda e acqua fredda |
| `Tabella_2026`                  | Le spese dell'anno e le letture dei contatori principali di energia elettrica, gas e acqua                             |
| `Tabelle millesimali`           | Subalterno, energia utile e millesimi di riscaldamento e di acqua calda di ogni appartamento                           |
| `Inquilini`                     | L'elenco degli appartamenti con il nome dell'inquilino                                                                 |
| `Dati Report`                   | Tutti i dati del file FC_report in una tabella con i filtri                                                            |

Il programma scrive nel file questi dati:

- dal file FC_report: le letture finali dei due contatori di calore della centrale termica, le letture finali di riscaldamento, acqua calda e acqua fredda di ogni appartamento, la data di lettura e i nomi degli appartamenti;
- dal file dei dati statici: i nomi degli inquilini, i subalterni e l'energia utile, le spese dell'anno, le letture dei contatori principali e le letture precedenti di ogni appartamento.

Quasi tutti i calcoli sono formule di Excel già presenti nel file: se correggi un dato a mano, i risultati si aggiornano. Fanno eccezione le celle elencate qui sotto.

### Dati da controllare a mano

Il programma non aggiorna questi dati. Controllali ogni volta nel primo foglio:

- le letture di inizio esercizio dei due contatori di calore della centrale termica, nelle celle `C27` e `C30`;
- le date di inizio esercizio, nelle celle `D27`, `D30` e `B178`;
- le quote delle spese fisse di riscaldamento e di acqua calda, nelle celle `C63` e `C64`. Sono numeri fissi, non formule. `C63` deve contenere lo stesso numero della cella `C18` del foglio `Tabelle millesimali`, e `C64` lo stesso numero della cella `E18` di quel foglio. Se sono diversi, scrivi in `C63` e `C64` i numeri che leggi in `C18` e `E18`.

### Se non usi il file dei dati statici

Il file creato contiene solo le letture Siemens. Mancano le spese e i millesimi, quindi molte celle mostrano `#DIV/0!`. Le letture precedenti degli appartamenti restano quelle già presenti nel modello e possono non essere le tue.

Hai due possibilità:

- prepara il file dei dati statici e ripeti la conversione con i due file;
- scrivi i dati mancanti a mano nel file creato.

## Il file dei dati statici

Il file dei dati statici contiene i dati che il sistema Siemens non conosce. Lo compili una volta e lo aggiorni ogni anno.

### Scaricare il modello

1. Apri la [pagina dell'ultima versione](https://github.com/gjed/siemens-converter/releases/latest).
1. Nell'elenco `Assets` clicca su `static_data_template.xlsx`.
1. Salva il file con un nome a tua scelta, per esempio `dati_statici_2026.xlsx`. Il nome deve finire con `.xlsx`.
1. Apri il file con Excel.
1. Compila i cinque fogli come spiegato qui sotto.
1. Salva il file.
1. Chiudi Excel.

### Regole per compilarlo

- Non cambiare il nome dei fogli.
- Non spostare, aggiungere o togliere colonne.
- Non toccare le prime due righe di ogni foglio: contengono il titolo e le intestazioni.
- Non cambiare i numeri degli appartamenti nella colonna `App.`.
- Nelle celle dei numeri scrivi solo il numero, senza unità di misura e senza il simbolo `€`.
- Non lasciare vuota una cella che deve contenere un numero. Il programma legge una cella vuota come zero e non ti avvisa. Se un valore è davvero zero, scrivi `0`.
- Non lasciare righe vuote in mezzo a una tabella. Il programma può smettere di leggere alla prima riga vuota.
- Non scrivere nelle celle che contengono già una formula.

### Foglio `Inquilini`

Una riga per ogni appartamento.

| Colonna                    | Cosa scrivere                                                                             |
| -------------------------- | ----------------------------------------------------------------------------------------- |
| `App.`                     | Il numero dell'appartamento, da 1 a 10. È già compilato                                   |
| `Proprietario / Condomino` | Il nome del proprietario. Oggi questo dato non viene riportato nel file creato            |
| `Inquilino / Conduttore`   | Il nome dell'inquilino. Compare nel foglio `Inquilini` e accanto alle letture del riparto |

I nomi degli appartamenti nel file creato non vengono da questo foglio: sono le descrizioni scritte nel file FC_report.

### Foglio `Millesimali`

Una riga per ogni appartamento.

| Colonna                       | Cosa scrivere                                       |
| ----------------------------- | --------------------------------------------------- |
| `App.`                        | Il numero dell'appartamento. È già compilato        |
| `Subal-terno`                 | Il numero del subalterno, senza decimali            |
| `Energia utile riscaldamento` | L'energia utile per il riscaldamento, in kWh        |
| `Millesimi risc.`             | Niente: è una formula                               |
| `Energia utile ACS`           | L'energia utile per l'acqua calda sanitaria, in kWh |
| `Millesimi ACS`               | Niente: è una formula                               |

Anche la riga `TOTALE` contiene formule: non modificarla.

### Foglio `Costi`

Le sei spese dell'anno. Scrivi l'importo in euro nella colonna `€ / anno`, accanto a ogni voce:

- `Energia elettrica`
- `Gas metano`
- `Acqua condominio`
- `Conduzione e manutenzione`
- `Contabilizzazione`
- `Acqua sanitaria manutenzione`

Non modificare e non aggiungere voci nella colonna `Voce di spesa`: il programma accetta solo queste sei e si ferma con un errore se ne trova un'altra. La riga `TOTALE` è una formula.

### Foglio `Contatori`

Le letture dei tre contatori principali: `ENERGIA ELETTRICA CT`, `GAS METANO CT` e `ACQUA GENERALE`.

| Colonna            | Cosa scrivere                                    |
| ------------------ | ------------------------------------------------ |
| `Contatore`        | Niente: il nome è già compilato. Non modificarlo |
| `Unità`            | Niente: è già compilata                          |
| `Lettura iniziale` | La lettura del contatore a inizio anno           |
| `Lettura finale`   | La lettura del contatore a fine anno             |
| `Consumo`          | Niente: è una formula                            |

Se cambi il nome di un contatore, il programma ignora quella riga senza avvisarti.

### Foglio `Letture_precedenti`

Una riga per ogni appartamento, con le letture finali dell'anno precedente. Sono le letture iniziali del nuovo riparto.

| Colonna                       | Cosa scrivere                                                    |
| ----------------------------- | ---------------------------------------------------------------- |
| `App.`                        | Il numero dell'appartamento. È già compilato                     |
| `Riscaldamento lettura prec.` | La lettura precedente del riscaldamento, in kWh, senza decimali  |
| `ACS lettura prec.`           | La lettura precedente dell'acqua calda sanitaria, in metri cubi  |
| `AFS lettura prec.`           | La lettura precedente dell'acqua fredda sanitaria, in metri cubi |

## Problemi e soluzioni

Quando qualcosa non va, il programma mostra una finestra con un messaggio. Alcuni messaggi sono in inglese. Cerca qui sotto il testo che vedi.

### `File not found:` seguito dal nome di un file

Il programma non trova il file FC_report.

1. Controlla che il file esista ancora e che non sia stato spostato o rinominato.
1. Ripeti la conversione.

### `Static data file not found:` seguito dal nome di un file

Il programma non trova il file dei dati statici.

1. Controlla che il file esista ancora e che non sia stato spostato o rinominato.
1. Ripeti la conversione.

### `Errore nel file dati statici: Static data file is missing required sheet(s):` seguito da uno o più nomi

Nel file dei dati statici mancano i fogli elencati nel messaggio. Succede se un foglio è stato rinominato o eliminato, oppure se hai scelto un file Excel che non è il file dei dati statici.

1. Controlla di aver scelto il file giusto.
1. Apri il file con Excel.
1. Controlla che i fogli si chiamino esattamente `Inquilini`, `Millesimali`, `Costi`, `Contatori` e `Letture_precedenti`.
1. Se un foglio manca, scarica di nuovo il modello e ricopia i tuoi dati.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici: Unknown cost label(s) in Costi:` seguito da una o più voci

Nel foglio `Costi` c'è una voce di spesa che il programma non conosce. Il messaggio la riporta in lettere minuscole.

1. Apri il file dei dati statici con Excel.
1. Vai al foglio `Costi`.
1. Nella colonna `Voce di spesa` cerca la voce riportata nel messaggio.
1. Correggila con uno dei sei nomi elencati in [Foglio `Costi`](#foglio-costi), oppure elimina la riga.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un altro testo in inglese

Il programma non riesce a leggere il file dei dati statici. Due esempi:

- `could not convert string to float: 'abc'`: una cella che deve contenere un numero contiene del testo. Il testo tra apici è quello trovato nella cella.
- `File is not a zip file`: il file non è un vero file Excel `.xlsx`.

1. Apri il file dei dati statici con Excel.
1. Controlla che le celle dei numeri contengano solo numeri.
1. Scegli `File`, poi `Salva con nome`, e salva come `Cartella di lavoro di Excel (*.xlsx)`.
1. Chiudi Excel e ripeti la conversione.

### `Impossibile aprire la finestra di selezione file:` seguito da un testo

Dopo il doppio clic il programma non riesce ad aprire la finestra per scegliere il file.

1. Premi `OK`.
1. Converti [trascinando i file](#convertire-trascinando-i-file) sopra l'icona del programma.

### Una finestra dal titolo `siemens-converter - Errore` con un lungo testo in inglese

Il programma ha incontrato un errore che non sa spiegare. Le cause più frequenti sono due.

- **Il file trascinato non è un file FC_report.** Succede, per esempio, se trascini solo il file dei dati statici: in quel caso il testo finisce con una riga che comincia con `UnicodeDecodeError`. Ripeti la conversione con il file FC_report esportato dal sistema Siemens.
- **Il file `Riparto` di una conversione precedente è aperto in Excel.** Chiudi quel file e ripeti la conversione.

Se il problema resta, fai una foto della finestra e inviala a chi ti ha fornito il programma.

## Per sviluppatori: riga di comando

Serve Python 3.15 con il pacchetto installato (vedi la [guida per sviluppatori](development.md)).

```bash
python -m siemens_converter FC_report.xls [static_data.xlsx]
```

- Il primo argomento è il file FC_report ed è obbligatorio.
- Il secondo argomento è il file dei dati statici ed è facoltativo.
- L'ordine dei due argomenti non conta: il file `.xlsx` viene usato come file dei dati statici.
- Il file `Riparto_<numero di serie>_<data>.xlsx` viene creato nella cartella del file FC_report.
- Senza argomenti il comando stampa le istruzioni d'uso ed esce con codice 1.
- In caso di errore il comando stampa il messaggio nel terminale ed esce con codice 1.
