# Guida all'uso

Il convertitore legge il file delle letture esportato dal sistema Siemens (il file FC_report) e crea un file Excel con la ripartizione delle spese di riscaldamento e acqua tra gli appartamenti.

## Cosa ti serve

- **Il programma.** È un file `.exe` con un nome simile a `siemens-converter-windows-v0.2.0.exe`. Non va installato.
- **Il file FC_report.** È il file `.xls` esportato dal sistema Siemens. Il suo nome comincia con `FC_report`.
- **Il file dei dati statici.** È un file Excel `.xlsx` che compili tu con l'anno dell'esercizio, inquilini, millesimi, costi e letture di inizio esercizio. Serve sempre: senza, il programma non converte. Come prepararlo è spiegato in [Il file dei dati statici](#il-file-dei-dati-statici).

## Scaricare il programma

1. Apri la [pagina dell'ultima versione](https://github.com/gjed/siemens-converter/releases/latest).
1. Nell'elenco `Assets` clicca sul file che termina con `.exe`.
1. Salva il file in una cartella che ritrovi facilmente, per esempio il Desktop.

## Convertire trascinando i file

### Il file FC_report insieme al file dei dati statici

1. Metti il file FC_report e il file dei dati statici nella stessa cartella.
1. Clicca una volta sul file FC_report.
1. Tieni premuto il tasto `Ctrl` e clicca una volta sul file dei dati statici. Ora i due file sono selezionati.
1. Trascina i due file insieme sopra l'icona del programma.
1. Rilascia i file.
1. Aspetta la finestra con la scritta `Fatto!` e il nome del file creato.
1. Premi `OK`.

Non importa quale file selezioni per primo. Il programma li riconosce dall'estensione: il file `.xlsx` è quello dei dati statici, l'altro è il file FC_report. Servono sempre tutti e due. Se ne trascini uno solo, o se i due file sono entrambi `.xlsx` o nessuno dei due lo è, il programma lo dice e non converte nulla.

Trascina al massimo due file. Il programma ignora quelli in più.

## Convertire con il doppio clic

1. Fai doppio clic sull'icona del programma.
1. Si apre la finestra `Seleziona FC_report (.xls)`. Clicca sul file FC_report.
1. Premi `Apri`.
1. Si apre la finestra `Seleziona dati statici (.xlsx)`. Clicca sul file dei dati statici.
1. Premi `Apri`.
1. Aspetta la finestra con la scritta `Fatto!` e il nome del file creato.
1. Premi `OK`.

Se premi `Annulla` nella prima finestra, il programma si chiude senza creare nulla. Se premi `Annulla` nella seconda, il programma dice che manca il file dei dati statici e non crea nulla.

## Dove trovi il risultato

Il file creato si trova nella stessa cartella del file FC_report.

Il suo nome è `Riparto_<numero di serie>_<data>.xlsx`, per esempio `Riparto_TE99999999_15-01-2026.xlsx`. Il numero di serie e la data (giorno-mese-anno) sono quelli scritti dentro il file FC_report.

Se nella cartella c'è già un file con lo stesso nome, il programma lo sostituisce senza chiedere conferma. Se avevi modificato a mano quel file, prima salvane una copia con un altro nome.

Se il programma trova qualcosa da controllare, crea comunque il file. Gli avvisi non fermano la conversione: compaiono nella finestra finale, sotto la scritta `Attenzione:`, e nel terminale prima della riga `Fatto:`. Che cosa significano e che cosa fare è spiegato in [Problemi e soluzioni](#problemi-e-soluzioni).

## Cosa contiene il file creato

Il file ha cinque fogli.

| Foglio                | Contenuto                                                                                                              |
| --------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| `Ripartizione`        | Il calcolo della ripartizione: letture iniziali e finali, consumi e costi di riscaldamento, acqua calda e acqua fredda |
| `Tabella`             | Le spese dell'anno e le letture dei contatori principali di energia elettrica, gas e acqua                             |
| `Tabelle millesimali` | Subalterno, energia utile e millesimi di riscaldamento e di acqua calda di ogni appartamento                           |
| `Inquilini`           | L'elenco degli appartamenti con il nome del proprietario e dell'inquilino                                              |
| `Dati Report`         | Tutti i dati del file FC_report in una tabella con i filtri                                                            |

I fogli `Ripartizione` e `Tabella` hanno sempre questi nomi, qualunque sia l'anno.

Il programma scrive nel file questi dati:

- dal file FC_report: le letture finali dei due contatori di calore della centrale termica, le letture finali di riscaldamento, acqua calda e acqua fredda di ogni appartamento, la data di lettura e i nomi degli appartamenti;
- dal file dei dati statici: l'anno e la data di inizio dell'esercizio, le letture di inizio esercizio dei due contatori di calore della centrale termica, i nomi dei proprietari e degli inquilini, i subalterni e l'energia utile, le spese dell'anno, le letture dei contatori principali e le letture precedenti di ogni appartamento. Se nel file dei dati statici scrivi la lettura finale di un contatore di calore della centrale termica, con la sua data, il programma usa quella al posto della lettura del file FC_report.

Quasi tutti i calcoli sono formule di Excel già presenti nel file: se correggi un dato a mano, i risultati si aggiornano. Fanno eccezione le celle elencate qui sotto.

### Dati da scrivere a mano

Il programma non conosce i giorni di permanenza degli inquilini. Se un inquilino è entrato in corso d'anno, scrivili a mano nel primo foglio, nelle righe da 257 a 273: i giorni di permanenza dell'inquilino nella colonna `N` e una nota, per esempio la data di entrata, nella colonna `K`. Se nessun inquilino è entrato in corso d'anno, non scrivere nulla.

## Il file dei dati statici

Il file dei dati statici contiene i dati che il sistema Siemens non conosce. Lo prepari ogni anno: le letture di inizio esercizio sono le letture finali dell'anno prima.

### Scaricare il modello

1. Apri la [pagina dell'ultima versione](https://github.com/gjed/siemens-converter/releases/latest).
1. Nell'elenco `Assets` clicca su `static_data_template.xlsx`.
1. Salva il file con un nome a tua scelta, per esempio `dati_statici_2026.xlsx`. Il nome deve finire con `.xlsx`.
1. Apri il file con Excel.
1. Compila i sei fogli come spiegato qui sotto.
1. Salva il file.
1. Chiudi Excel.

### Regole per compilarlo

- Non cambiare il nome dei fogli.
- Non spostare, aggiungere o togliere colonne.
- Non toccare le prime due righe di ogni foglio: contengono il titolo e le intestazioni.
- Non cambiare i numeri degli appartamenti nella colonna `App.`.
- Nelle celle dei numeri scrivi solo il numero, senza unità di misura e senza il simbolo `€`.
- Non lasciare vuota una cella che deve contenere un numero. Il programma si ferma e ti dice quale cella manca. Se un valore è davvero zero, scrivi `0`.
- Scrivi ogni appartamento una sola volta, da 1 a 10, in ogni foglio che ha la colonna `App.`. Non cancellare righe e non scrivere lo stesso numero in due righe.
- Non scrivere note, totali o altre righe sotto o in mezzo alle tabelle: oltre alle righe dei dati e alla riga `TOTALE`, il programma non accetta altro e si ferma.
- Non scrivere nelle celle che contengono già una formula.

### Foglio `Inquilini`

Una riga per ogni appartamento.

| Colonna                    | Cosa scrivere                                                                              |
| -------------------------- | ------------------------------------------------------------------------------------------ |
| `App.`                     | Il numero dell'appartamento, da 1 a 10. È già compilato                                    |
| `Proprietario / Condomino` | Il nome del proprietario. Compare nella colonna `A` del foglio `Inquilini` del file creato |
| `Inquilino / Conduttore`   | Il nome dell'inquilino. Compare nel foglio `Inquilini` e accanto alle letture del riparto  |

Se non scrivi il proprietario di un appartamento, nel file creato compare la descrizione scritta nel file FC_report, per esempio `App, 01 Rossi Mario`.

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

Le righe dei dati cominciano alla riga 5, sotto le unità di misura e la riga `TOTALE`, che contiene formule: non modificarla.

### Foglio `Costi`

Le sei spese dell'anno. Scrivi l'importo in euro nella colonna `€ / anno`, accanto a ogni voce:

- `Energia elettrica`
- `Gas metano`
- `Acqua condominio`
- `Conduzione e manutenzione`
- `Contabilizzazione`
- `Acqua sanitaria manutenzione`

Non modificare e non aggiungere voci nella colonna `Voce di spesa`: il programma accetta solo queste sei, ognuna una volta, e si ferma con un errore se ne trova un'altra o se ne manca una. Se una spesa è zero, scrivi `0`. La riga `TOTALE` è una formula.

### Foglio `Contatori`

Le letture dei tre contatori principali: `ENERGIA ELETTRICA CT`, `GAS METANO CT` e `ACQUA GENERALE`.

| Colonna            | Cosa scrivere                                    |
| ------------------ | ------------------------------------------------ |
| `Contatore`        | Niente: il nome è già compilato. Non modificarlo |
| `Unità`            | Niente: è già compilata                          |
| `Lettura iniziale` | La lettura del contatore a inizio anno           |
| `Lettura finale`   | La lettura del contatore a fine anno             |
| `Consumo`          | Niente: è una formula                            |

Se cambi il nome di un contatore, o ne togli uno, il programma si ferma con un errore. Si ferma anche se la lettura finale è più bassa di quella iniziale.

### Foglio `Letture_precedenti`

Una riga per ogni appartamento, con le letture finali dell'anno precedente. Sono le letture iniziali del nuovo riparto.

| Colonna                       | Cosa scrivere                                                    |
| ----------------------------- | ---------------------------------------------------------------- |
| `App.`                        | Il numero dell'appartamento. È già compilato                     |
| `Riscaldamento lettura prec.` | La lettura precedente del riscaldamento, in kWh, senza decimali  |
| `ACS lettura prec.`           | La lettura precedente dell'acqua calda sanitaria, in metri cubi  |
| `AFS lettura prec.`           | La lettura precedente dell'acqua fredda sanitaria, in metri cubi |

Il programma controlla che le letture finali del file FC_report non siano più basse di queste. Se lo sono, il file dei dati statici è di un altro anno e il programma si ferma.

### Foglio `Esercizio`

L'anno, la data di inizio e le letture dei due contatori di calore della centrale termica.

| Riga                                  | Colonna `Valore`                                         | Colonne facoltative                      |
| ------------------------------------- | -------------------------------------------------------- | ---------------------------------------- |
| `Anno dell'esercizio`                 | L'anno con quattro cifre, per esempio `2026`             | Niente                                   |
| `Data di inizio dell'esercizio`       | La data delle letture iniziali, per esempio `31/12/2025` | Niente                                   |
| `Contacalorie centrale RISCALDAMENTO` | La lettura iniziale, in kWh                              | `Lettura finale` e `Data lettura finale` |
| `Contacalorie centrale ACS`           | La lettura iniziale, in kWh                              | `Lettura finale` e `Data lettura finale` |

L'anno compare nelle etichette `FINE GESTIONE` del primo foglio. La data di inizio compare accanto alle letture iniziali.

I contatori di calore della centrale termica spesso sono letti qualche giorno prima di quelli degli appartamenti. Se vuoi usare una lettura diversa da quella del file FC_report, scrivi la lettura finale e la sua data nelle colonne facoltative: servono tutte e due. Se le lasci vuote, il programma usa la lettura del file FC_report. Gli avvisi sulla data di lettura e sullo stato del contatore centrale riguardano la lettura che il programma usa: se hai scritto la tua, valgono la tua data e il suo valore, non quelli del file FC_report.

## Problemi e soluzioni

Quando qualcosa non va, il programma mostra una finestra con un messaggio scritto in italiano. Il messaggio dice che cosa non va e che cosa fare. Cerca qui sotto l'inizio del testo che vedi.

### `Manca il file dei dati statici: è stato scelto solo` seguito dal nome di un file

Hai scelto solo il file FC_report, oppure hai premuto `Annulla` nella seconda finestra del doppio clic. Il programma non converte senza il file dei dati statici, perché senza di esso il riparto sarebbe sbagliato.

1. Prepara il file dei dati statici, come spiegato in [Il file dei dati statici](#il-file-dei-dati-statici).
1. Scegli il file FC_report insieme al file dei dati statici, come spiegato in [Il file FC_report insieme al file dei dati statici](#il-file-fc_report-insieme-al-file-dei-dati-statici).
1. Ripeti la conversione.

### `Manca il file FC_report: è stato scelto solo` seguito dal nome di un file

Hai scelto un solo file e il suo nome finisce con `.xlsx`. Il programma lo considera il file dei dati statici e non trova il file FC_report.

1. Riapri la cartella che contiene il file FC_report.
1. Scegli il file FC_report insieme al file dei dati statici, come spiegato in [Il file FC_report insieme al file dei dati statici](#il-file-fc_report-insieme-al-file-dei-dati-statici).
1. Ripeti la conversione.

### `Sono stati scelti due file .xlsx`

Hai scelto due file che finiscono con `.xlsx`. Il programma ne accetta uno solo, il file dei dati statici, e non trova il file FC_report.

1. Togli il file `.xlsx` in più.
1. Scegli il file FC_report insieme al file dei dati statici.
1. Ripeti la conversione.

### `Sono stati scelti due file che non sono .xlsx`

Hai scelto due file che non finiscono con `.xlsx`. Il programma accetta un solo file FC_report e vuole che il file dei dati statici sia un file `.xlsx`.

1. Scegli il file FC_report e il file dei dati statici. Il nome del file dei dati statici deve finire con `.xlsx`.
1. Ripeti la conversione.

### `Il file FC_report` seguito da un nome e da `non esiste`

Il programma non trova il file FC_report.

1. Controlla che il file esista ancora e che non sia stato spostato o rinominato.
1. Ripeti la conversione.

### `Il file dei dati statici` seguito da un nome e da `non esiste`

Il programma non trova il file dei dati statici.

1. Controlla che il file esista ancora e che non sia stato spostato o rinominato.
1. Ripeti la conversione.

### `Errore nel report FC: il file` seguito da un nome e da `non è un report FC`

Il file scelto come FC_report non è un file FC_report. Dopo `non è un report FC:` il messaggio dice perché. Per esempio:

- `non contiene testo leggibile`: il file non è un file di testo. Succede se hai scelto un file Excel, un'immagine o un PDF.
- `non contiene nessuna tabella`: il file è un testo, ma non è un report del sistema Siemens.
- `la tabella è troppo corta` o `la tabella ha meno colonne del previsto`: il file è stato modificato o salvato in un altro formato, oppure è incompleto.
- `il numero di serie` o `il numero dei dispositivi collegati` `non è un numero`: il file è stato modificato a mano.

1. Controlla di aver scelto il file giusto. Il nome del file FC_report comincia con `FC_report` e finisce con `.xls`.
1. Non aprire e non salvare di nuovo il file con Excel prima di convertirlo: Excel lo cambia.
1. Se il file è quello giusto, scaricalo di nuovo dal sistema Siemens.
1. Ripeti la conversione.

### `Errore nel report FC:` seguito da un dispositivo e da `nella descrizione manca il numero dell'appartamento`

Nel file FC_report la descrizione di un dispositivo non contiene il numero dell'appartamento. Il messaggio riporta la descrizione trovata. Il programma si aspetta una descrizione che cominci come `App, 01`.

1. Fai correggere la descrizione del dispositivo nel sistema Siemens da chi lo gestisce.
1. Scarica di nuovo il file FC_report.
1. Ripeti la conversione.

### `Errore nel report FC: il file` seguito da un nome e da `non si può aprire per leggerlo`

Il programma non riesce ad aprire il file FC_report. Succede se hai scelto una cartella invece di un file, se il file è aperto in un altro programma o se non hai il permesso di leggerlo.

1. Controlla di aver scelto un file e non una cartella.
1. Chiudi il file negli altri programmi che lo tengono aperto.
1. Se il file è su una chiavetta o su un'unità di rete, copialo sul Desktop.
1. Ripeti la conversione.

### `Errore nel report FC: il report non è completo e il riparto non è stato creato.`

Il programma non ha creato il file `Riparto`, perché con questi dati il risultato sarebbe sbagliato. Il messaggio elenca che cosa non va in una o due liste.

#### `Dispositivi mancanti:`

Nel file FC_report non ci sono i dispositivi elencati. Ogni riga dice quale dispositivo manca: un contatore centrale (`Riscaldamento` o `Sanitario`), oppure un contacalorie o un contatore dell'acqua calda di un appartamento.

1. Scarica di nuovo il file FC_report dal sistema Siemens.
1. Ripeti la conversione.
1. Se il messaggio compare ancora, fai controllare sull'impianto i dispositivi elencati da chi gestisce il sistema Siemens. Controlla che siano installati, alimentati e collegati.
1. Quando i dispositivi tornano nel report, ripeti la conversione.

#### `Letture mancanti o non valide:`

Nel file FC_report c'è un dispositivo, ma una lettura che serve è vuota o non è un numero. Ogni riga dice quale dispositivo e quale lettura: energia, acqua calda o acqua fredda. Se la lettura non è un numero, la riga riporta il testo trovato.

1. Scarica di nuovo il file FC_report dal sistema Siemens.
1. Ripeti la conversione.
1. Se il messaggio compare ancora, il dispositivo non ha trasmesso la lettura. Fallo controllare sull'impianto da chi gestisce il sistema Siemens.
1. Quando il dispositivo trasmette la lettura, ripeti la conversione.

### `Attenzione:` nella finestra `Fatto!`

Il programma ha creato il file `Riparto`, ma ha trovato qualcosa da controllare. Ogni riga dopo `Attenzione:` è un avviso. Gli avvisi non fermano la conversione: leggili tutti e controlla il file prima di usarlo. Gli avvisi sono tre.

#### `Il contatore centrale` seguito da `è stato letto il`

Il contatore centrale è stato letto in un giorno diverso da quello dei contatori degli appartamenti. Il messaggio dice quanti giorni di differenza ci sono e le due date. I consumi del contatore centrale e quelli degli appartamenti non coprono lo stesso periodo.

1. Apri il file creato e guarda il primo foglio. Le date di lettura dei contatori centrali sono nelle celle `D28` e `D31`, quelle degli appartamenti nelle celle `C76`, `C126` e `C178`.
1. Se la differenza è di pochi giorni, decidi insieme all'amministratore se il riparto si può usare.
1. Altrimenti fai leggere di nuovo tutti i contatori nello stesso giorno dal sistema Siemens, scarica un nuovo file FC_report e ripeti la conversione.

#### `I contacalorie degli appartamenti non sono stati letti tutti lo stesso giorno`

Lo stesso avviso esiste per `I contatori dell'acqua calda degli appartamenti`. Alcuni contatori degli appartamenti sono stati letti prima degli altri. Il messaggio elenca i contatori letti prima, con la loro data. Nel file creato c'è la data più recente.

1. Apri il file creato e controlla nel primo foglio le letture dei contatori elencati. Una lettura fatta qualche giorno prima può essere più bassa di quella del giorno più recente.
1. Se la differenza è di pochi giorni, decidi insieme all'amministratore se il riparto si può usare.
1. Altrimenti fai leggere di nuovo i contatori elencati dal sistema Siemens, scarica un nuovo file FC_report e ripeti la conversione.

#### `Lo stato di questi dispositivi nel report non è «Ok»`

Il messaggio elenca i dispositivi che il sistema Siemens segnala con uno stato diverso da `Ok`, e tra parentesi lo stato. Se tra parentesi c'è `vuoto`, il report non indica nessuno stato. Le letture di questi dispositivi possono essere sbagliate.

1. Apri il file creato e controlla nel primo foglio le letture di questi dispositivi. Confrontale con le letture precedenti e con quelle che leggi sul contatore.
1. Se una lettura non è credibile, fai controllare il dispositivo sull'impianto da chi gestisce il sistema Siemens.
1. Dopo il controllo scarica un nuovo file FC_report e ripeti la conversione. In alternativa correggi a mano la lettura nel file creato.

### `Errore nel file dati statici:` seguito da un nome e da `non è un file dei dati statici`

Il file scelto come file dei dati statici non è un file dei dati statici. Il messaggio dice perché:

- `non si apre come file Excel (.xlsx)`: il file non è un vero file Excel `.xlsx`. Succede se hai scelto un altro tipo di file o se hai solo cambiato l'estensione di un file.
- `manca il foglio` o `mancano i fogli`, seguito dai nomi: nel file mancano fogli del modello. Succede se un foglio è stato rinominato o eliminato.

1. Controlla di aver scelto il file giusto.
1. Apri il file con Excel.
1. Controlla che i fogli si chiamino esattamente `Inquilini`, `Millesimali`, `Costi`, `Contatori`, `Letture_precedenti` ed `Esercizio`.
1. Se un foglio manca, scarica di nuovo il modello come spiegato in [Scaricare il modello](#scaricare-il-modello) e ricopia i tuoi dati. Se manca il foglio `Esercizio`, il file è stato fatto con un modello vecchio: scarica il modello nuovo.
1. Salva il file come `Cartella di lavoro di Excel (*.xlsx)`, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da `foglio «Costi»` o `foglio «Contatori»` e da `non sono tra quelle accettate` o `non sono tra quelli accettati`

Nel foglio c'è una voce di spesa o un contatore che il programma non conosce. Il messaggio riporta ogni nome come l'hai scritto, con la cella, e elenca quelli accettati.

1. Apri il file dei dati statici con Excel.
1. Vai al foglio indicato.
1. Cerca il nome nella cella indicata dal messaggio.
1. Correggilo con uno dei nomi elencati, scritto come nel modello (vedi [Foglio `Costi`](#foglio-costi) e [Foglio `Contatori`](#foglio-contatori)), oppure elimina la riga.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da un foglio e da `manca la voce`, `mancano le voci`, `manca il contatore` o `mancano i contatori`

Nel foglio `Costi` o `Contatori` manca una riga. Il messaggio dice quale. Il programma non la considera zero: ogni voce e ogni contatore devono esserci.

1. Apri il file dei dati statici con Excel.
1. Aggiungi la riga che manca, scritta come nel modello, con i suoi valori. Se un valore è davvero zero, scrivi `0`.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da un foglio e da `compare due volte`

Nel foglio c'è due volte lo stesso appartamento, la stessa voce di spesa o lo stesso contatore. Il messaggio dice quale e in quali celle.

1. Apri il file dei dati statici con Excel.
1. Vai al foglio e alle celle indicati.
1. Lascia una sola riga, quella giusta, e correggi o elimina l'altra.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da un foglio e da `manca l'appartamento` o `mancano gli appartamenti`

Nel foglio `Inquilini`, `Millesimali` o `Letture_precedenti` non c'è la riga di uno o più appartamenti. Il programma vuole tutti gli appartamenti da 1 a 10.

1. Apri il file dei dati statici con Excel.
1. Controlla nel foglio indicato che ci sia una riga per ogni appartamento da 1 a 10.
1. Aggiungi la riga che manca, con il numero nella colonna `App.`, e compilala.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da un foglio, da una cella e da `non è un numero di appartamento` o `non esiste`

Nella colonna `App.` c'è un valore che non è il numero di un appartamento da 1 a 10: per esempio una nota, un totale scritto a mano o il numero 11. Il programma accetta solo le righe degli appartamenti e la riga `TOTALE`.

1. Apri il file dei dati statici con Excel.
1. Vai alla cella indicata dal messaggio.
1. Correggi il numero oppure cancella tutta la riga.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da una cella e da `manca il numero dell'appartamento` o `manca il nome`

Una riga ha dei valori, ma la cella con il numero dell'appartamento (o con il nome della voce o del contatore) è vuota.

1. Apri il file dei dati statici con Excel.
1. Vai alla cella indicata dal messaggio.
1. Scrivi il numero o il nome mancante, come nel modello, oppure cancella tutta la riga.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da un foglio, da una cella e da `non è un numero`

Una cella che deve contenere un numero contiene del testo. Il messaggio dice il foglio, la cella e il valore trovato. Se dice `non è un numero intero`, la cella deve contenere un numero senza decimali.

1. Apri il file dei dati statici con Excel.
1. Vai al foglio e alla cella indicati.
1. Scrivi solo il numero, senza unità di misura e senza il simbolo `€`.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, da un foglio, da una cella e da `la cella è vuota ma serve un numero`

Una cella che deve contenere un numero è vuota. Il programma non la considera zero, perché un numero dimenticato sposterebbe dei soldi da un appartamento a un altro senza che nessuno se ne accorga.

1. Apri il file dei dati statici con Excel.
1. Vai al foglio e alla cella indicati.
1. Scrivi il numero. Se è davvero zero, scrivi `0`.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, dal foglio `Contatori` e da `la lettura finale` e `è più bassa della lettura iniziale`

Nel foglio `Contatori` la lettura finale di un contatore è più bassa della lettura iniziale, per esempio perché manca una cifra. Il messaggio dice il contatore, le due letture e la cella della lettura finale. Se le due letture sono uguali il programma le accetta.

1. Apri il file dei dati statici con Excel.
1. Vai al foglio `Contatori` e controlla le due letture del contatore indicato.
1. Correggi la lettura sbagliata.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Errore nel file dati statici:` seguito da un nome, dal foglio `Esercizio` e da `anno`, `data` o `lettura finale`

Nel foglio `Esercizio` manca un dato o non si legge. Il messaggio dice la cella.

- `l'anno dell'esercizio è vuoto` o `non è un anno`: scrivi l'anno con quattro cifre, per esempio `2026`.
- `la data di inizio dell'esercizio è vuota` o `non è una data`: scrivi la data come `31/12/2025`.
- `manca la data della lettura finale` o `manca la lettura finale della data`: la lettura finale e la sua data vanno scritte insieme. Scrivile tutte e due oppure cancella quella che hai scritto.

1. Apri il file dei dati statici con Excel.
1. Vai alla cella indicata dal messaggio e correggila.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Alcune letture finali sono più basse delle letture iniziali dei dati statici`

Il file dei dati statici e il file FC_report non sono dello stesso esercizio, oppure una lettura è sbagliata. Il messaggio elenca ogni contatore con la lettura finale (del file FC_report, o scritta nel foglio `Esercizio` se l'hai usata) e la lettura iniziale (del file dei dati statici). Il programma non crea il file `Riparto`, perché i consumi sarebbero negativi.

1. Controlla di aver scelto il file dei dati statici dello stesso esercizio del file FC_report. Le letture iniziali devono essere le letture finali dell'esercizio precedente.
1. Se il file è quello giusto, apri il file dei dati statici con Excel e controlla le letture dei contatori elencati nei fogli `Letture_precedenti` ed `Esercizio`.
1. Salva il file, chiudi Excel e ripeti la conversione.

### `Impossibile creare il file` seguito da un nome e da una cartella

Il programma non riesce a scrivere il file `Riparto`. La causa più frequente è che il file `Riparto` di una conversione precedente è aperto in Excel. Un'altra è che non hai il permesso di scrivere nella cartella del file FC_report.

1. Chiudi in Excel il file `Riparto` indicato dal messaggio.
1. Se non era aperto, copia il file FC_report in una cartella in cui puoi scrivere, per esempio il Desktop.
1. Ripeti la conversione.

### `Impossibile aprire la finestra di selezione dei file.`

Dopo il doppio clic il programma non riesce ad aprire la finestra per scegliere il file.

1. Premi `OK`.
1. Converti [trascinando i file](#convertire-trascinando-i-file) sopra l'icona del programma.

### `La conversione non è riuscita per un errore imprevisto del programma.`

Questa frase apre una finestra dal titolo `siemens-converter - Errore`. Sotto la frase, dopo `Dettagli tecnici:`, c'è un testo tecnico in inglese. Il programma ha incontrato un errore che non sa spiegare.

1. Ripeti la conversione.
1. Se la finestra compare ancora, fai una foto della finestra, o copia il testo, e inviala a chi ti ha fornito il programma insieme al nome dei file che stavi convertendo.

## Per sviluppatori: riga di comando

Serve Python 3.14 o 3.15 con il pacchetto installato (vedi la [guida per sviluppatori](development.md)).

```bash
python -m siemens_converter FC_report.xls static_data.xlsx
```

- Il primo argomento è il file FC_report ed è obbligatorio.
- Il secondo argomento è il file dei dati statici ed è obbligatorio: con un solo file il comando stampa un messaggio ed esce con codice 1.
- L'ordine dei due argomenti non conta: il file `.xlsx` viene usato come file dei dati statici.
- Il file `Riparto_<numero di serie>_<data>.xlsx` viene creato nella cartella del file FC_report.
- Senza argomenti il comando stampa le istruzioni d'uso ed esce con codice 1.
- In caso di errore il comando stampa il messaggio nel terminale ed esce con codice 1.
