# Stesso bilancio su telefono e PC

I dati (spese, voci fisse) vengono salvati in un **archivio privato sul tuo GitHub**, che solo tu
puoi leggere. Ogni dispositivo si collega con un token. La app sincronizza da sola: all'apertura,
dopo ogni salvataggio e quando torna la rete. Poi **Crea l'Excel** esce identico dal telefono e dal PC.

## Preparazione (una volta sola, dal computer)

### 1. Crea l'archivio dei dati
1. Su github.com: **+ (in alto a destra) → New repository**.
2. Nome: `bilancio-dati`.
3. Seleziona **Private** (obbligatorio: sono i tuoi conti).
4. Spunta **Add a README file** (serve perché l'archivio non sia vuoto).
5. **Create repository**.

### 2. Crea il token (la "chiave" della app)
1. GitHub → tua foto in alto a destra → **Settings** → in fondo a sinistra **Developer settings**.
2. **Personal access tokens → Fine-grained tokens → Generate new token**.
3. **Token name:** `bilancio-casa`. **Expiration:** 1 anno (annotati la scadenza).
4. **Repository access → Only select repositories** → scegli `bilancio-dati`.
5. **Permissions → Repository permissions → Contents → Read and write**.
6. **Generate token** e **copia subito** il codice che inizia con `github_pat_`:
   GitHub lo mostra una volta sola. Incollalo in un appunto sicuro per poterlo usare sugli altri dispositivi.

## Collegare ogni dispositivo
Su **telefono** e su **PC**, nella app: scheda **Excel → Stesso bilancio su telefono e PC → Collega**.
- **Archivio:** `tuoutente/bilancio-dati`
- **Token:** quello copiato sopra

Premi **Collega e sincronizza**. Fatto: da ora le spese viaggiano da sole.

Sul **PC** apri nel browser (Chrome o Edge) lo stesso indirizzo `https://TUOUTENTE.github.io/Bilancio-casa/`.
Puoi anche installarla (icona di installazione nella barra dell'indirizzo).

## Aggiungere altri telefoni (famiglia)
1. Sul telefono già collegato: **Excel → Stesso bilancio su telefono e PC → Invita un altro telefono**.
2. Scegli WhatsApp, Messaggi o email e invia il link.
3. Sull'altro telefono apri il link (meglio con Safari su iPhone o Chrome su Android), poi **Collega**.
   Per averla come app: *Aggiungi alla schermata Home*, poi apri la app installata dal link ricevuto.

Il link contiene la chiave di accesso ai vostri dati: invialo solo a persone fidate.
Per togliere l'accesso a tutti, su GitHub elimina il token (Developer settings → Fine-grained tokens)
e creane uno nuovo da collegare ai telefoni che vuoi tenere.

## Condividere solo l'Excel
**Excel → Crea l'Excel**: finito il file compare **Condividi il file**. Scegli WhatsApp, email o Drive.
(Se il tuo telefono non lo mostra, il file è stato scaricato e lo trovi nella cartella Download o File.)

## Regole d'oro
- **Importa il tuo Excel una volta sola**, su un solo dispositivo: poi arriva sugli altri con la sincronizzazione.
  Se lo importi anche sull'altro, non crea doppioni.
- **Modifica ed elimina le spese dentro la app**, dal telefono o dal PC: la modifica arriva ovunque.
  L'Excel è il risultato: se lo cambi a mano, la app non lo sa. Reimportarlo aggiunge solo le righe nuove.
- **iPhone:** la app installata sulla Home ha una memoria separata da Safari. Se avevi già inserito
  spese in Safari prima di installarla, collega anche Safari, sincronizza, e poi collega la app installata:
  i dati arrivano da soli.
- **Senza rete:** puoi usare la app lo stesso. Appena torna la connessione, sincronizza.
- **Token scaduto o errato:** compare l'avviso "Token non valido o scaduto". Crea un token nuovo
  (passo 2) e ripeti **Collega**.
- **Copia di sicurezza:** resta utile: **Excel → Salva copia**. Il token non è dentro il file.

## Aggiornare la app su GitHub Pages
Nel repository `Bilancio-casa`: **Add file → Upload files**, trascina `app.js` (e `LEGGIMI_…` se vuoi),
**Commit changes**. I file con lo stesso nome vengono sostituiti. Alla seconda apertura della app
con internet, trovi la versione nuova.
