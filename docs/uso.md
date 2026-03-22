# Guida Utente

## Utilizzo con trascinamento (Windows)

1. Scarica il file `.exe` dalla [pagina release](https://github.com/gjed/siemens-converter/releases/latest)
1. Trascina il file CSV del report Siemens sull'eseguibile
1. Il file XLSX viene creato nella stessa cartella del CSV

## Utilizzo da riga di comando

```bash
python -m siemens_converter percorso/del/report.csv
```

## Formato di output

Il file XLSX generato contiene due fogli:

1. **PIVOT** -- riepilogo con i totali per appartamento
1. **Dati grezzi** -- tutti i dati del report originale con formattazione professionale

## Risoluzione problemi

- **"File not found"** -- verificare che il percorso del file CSV sia corretto
- **Caratteri speciali** -- il convertitore gestisce automaticamente i file con BOM e codifica UTF-8
