# Bilancio Casa: come metterla sul telefono

La app è una cartella di file. Per installarla sul telefono va messa online una sola volta
(gratis). Dopo la prima apertura funziona anche senza internet.
Gli scontrini non escono mai dal telefono: la lettura avviene lì.

## 1. Pubblicare la cartella (una volta sola, circa 10 minuti)

### Opzione A: GitHub Pages (gratuito, stabile)
1. Crea un account su github.com e poi un repository nuovo, per esempio `bilancio-casa`
   (visibilità *Public*; nessun dato personale è dentro la cartella).
2. Nel repository: **Add file → Upload files**. Trascina **tutto il contenuto** della cartella
   `BilancioCasa` (index.html, app.js, style.css, sw.js, manifest, e le cartelle `icons` e `vendor`).
   Se la cartella `vendor` non si carica in un colpo solo, caricala in due volte.
3. **Commit changes**.
4. **Settings → Pages → Build and deployment → Deploy from a branch → main / (root) → Save**.
5. Dopo 1-2 minuti compare l'indirizzo, del tipo `https://TUONOME.github.io/bilancio-casa/`.

### Opzione B: Netlify Drop (più rapida)
Vai su app.netlify.com/drop, trascina la cartella `BilancioCasa`, e crea l'accesso gratuito
per tenere il sito attivo. Ottieni subito un indirizzo.

## 2. Installarla sul telefono
- **iPhone (Safari):** apri l'indirizzo → tasto Condividi → **Aggiungi alla schermata Home**.
- **Android (Chrome):** apri l'indirizzo → menu ⋮ → **Installa app** (o *Aggiungi a schermata Home*).

Alla prima apertura, resta connesso per qualche secondo: la app scarica da sola il motore di
lettura (circa 12 MB). Poi è tutto offline.

## 3. Uso quotidiano
1. **Scansiona → Scatta una foto** (o *Scegli foto o PDF*, anche più file insieme).
2. Controlla importo e data, scegli **Chi**, la **Voce** (SPESA, BAR, FARMACIA…) e il **Tipo**
   (Ordinaria = colonna Importo, Extra = colonna Extra, Fissa = Importo + Fisso mensile). **Salva**.
3. **Excel → Crea l'Excel** quando ti serve il file: ottieni l'Excel nel formato FAMILIARE.


## 3b. Nomi e numero di persone
Nella scheda **Spese**, il pulsante **Nominativi** permette di cambiare i nomi, o di usare la app per
una sola persona (lascia vuoto il secondo nome). I nomi si aggiornano anche nelle spese già salvate,
nelle voci fisse e nell'Excel (SPESE - NOME), e arrivano da soli sugli altri dispositivi collegati.

## 4. Il tuo Excel (Fase 2)
- **Excel → Carica un Excel FAMILIARE**: carica un tuo file esistente (anche uno per anno). Le spese passano nell'app e si può ripetere senza doppioni.
- **Spese fisse e ricorrenti**: METLIFE, NETFLIX, ALLEANZA, NIMS - CAFFE, TELEFONO sono già
  inserite. Ogni mese compaiono da sole nell'Excel, fino al mese in corso. Importi, giorno e
  periodo si modificano toccando la voce.
- **Anno**: scegli *Gennaio – Dicembre* dell'anno che ti serve (2026, 2027, 2028…).
- L'ultimo foglio **RIEPILOGO** ha la tabella mensile (con formule collegate ai fogli) e due grafici.

## 5. Dati e sicurezza
- I dati stanno solo nel telefono. **Excel → Salva copia** crea un file `.json`: fallo ogni tanto,
  ed è anche il modo per portare i dati su un altro telefono (**Ripristina**).
- Se cancelli i dati del sito dal browser, le spese si perdono: per questo serve la copia.

## 6. Aggiornare la app in futuro
Sostituisci i file nel repository (o ritrascina la cartella su Netlify): la app si aggiorna
alla successiva apertura con internet.
