# HD Manager Desktop - Release v1.0.8

Applicazione desktop per la gestione di clienti, prodotti, risorse, VPN e archivio documentale.

---

## Note di rilascio v1.0.8

### Versione applicazione
- Versione **1.0.8**, data ultimo rilascio **15 aprile 2026** (da `app/version.py`: finestra principale, **Impostazioni → Informazione Prodotto**).
- Allineamento: `installer.iss`, `scripts/build_release.ps1`, queste note.

### Stabilità (Windows / eseguibile installato)
- Avvio Qt con percorsi corretti dal bundle PyInstaller (`app/qt_runtime.py`): `QT_PLUGIN_PATH`, plugin piattaforma e **`os.add_dll_directory`** sulla cartella `PyQt6/Qt6/bin` (riduce crash nativi in `Qt6Core.dll`, es. codice `0xc0000409`).
- Build release: PyInstaller con **`--collect-all PyQt6`** per includere DLL e plugin Qt necessari.
- Impostazioni: in chiusura finestra si attende il completamento del thread che legge l’elenco VPN Windows.

---

## Note di rilascio v1.0.7

### Versione applicazione
- Versione **1.0.7**, data ultimo rilascio **28 marzo 2026** (da `app/version.py`: finestra principale, **Impostazioni → Informazione Prodotto**).
- Allineamento: `installer.iss`, `scripts/build_release.ps1`, queste note.

### Avvio
- Finestra principale avviata **massimizzata** (area di lavoro).

### Archivio
- Pulsante **Apri cartella** nelle viste **Archivio** (schede File e Link), **Tag** e **Preferiti**: apre Esplora file sulla cartella del file selezionato (per link, solo percorsi locali / `file://`).

### Impostazioni
- **Tabelle**: editor e righe più alti, padding e stili per evitare testo tagliato in **QLineEdit** / **QComboBox** nelle celle.
- **Nessun refresh distruttivo** mentre una pagina è in **Modifica** (cambio voce di menu non ricarica dal DB se ci sono bozze).
- Completamento lista **VPN Windows** in background: in modalità modifica VPN non viene più eseguito un refresh completo che cancellava modifiche non salvate; aggiornati solo gli elenchi dei combo quando serve.

---

## Note di rilascio v1.0.6

### Versione applicazione
- Numero versione e data ultimo rilascio da `app/version.py` (anche nella finestra principale e in **Impostazioni → Informazione Prodotto**).
- Allineamento: `installer.iss`, `scripts/build_release.ps1`, queste note.

### Clienti — Rubrica
- Pulsanti **Chiama** (handler predefinito `tel:`; se presenti telefono e cellulare, scelta del numero) e **Invia Mail** (`mailto:`).

### Clienti — Contatti
- Pulsante **Programmare chiamata**: apre il dialog nuovo impegno in agenda con titolo precompilato «Chiamare …» sul nome del contatto.

### Impostazioni — Informazione Prodotto
- Nuova voce di menu con versione, data ultimo rilascio, autore (Diego Giotta), indicazione «Progetto Indipendente» e contatto e-mail cliccabile.

### Correzioni
- Etichetta e-mail nella pagina Informazione Prodotto: uso corretto di `Qt.TextInteractionFlag` (PyQt6).

---

## Note di rilascio v1.0.5

### Versione applicazione
- Titolo finestra principale con numero versione (da `app/version.py`).
- Allineamento versione: `installer.iss`, `scripts/build_release.ps1`, queste note.

### Impostazioni — Ruoli, risorse e competenze
- Vista unica con tre tabelle verticali: **Competenze** → **Ruoli** → **Risorse**.
- In **Risorse**: colonna opzionale **Competenza** (scelta dalle competenze definite).
- In **Info cliente**, sulla card persona: competenza mostrata sotto il nome se presente.

### SQL — Archivio query (linguetta SQL)
- **Gestione**: sostituzione **tabelle** (anche da `FROM` in sottoquery), **alias** (anche in `Alias.campo`), **solo nome campo** dopo il punto nelle espressioni qualificate.
- Sezioni a fisarmonica (**Tabelle** / **Alias di tabella** / **Campi qualificati**), stile coerente con il resto dell’app.
- Campi «Sostituisci con» **precompilati** con i valori rilevati (modificabili senza riscrivere tutto).
- Indicazione campi usati anche in WHERE/HAVING (tooltip).

### Installazione / aggiornamento
- Comportamento confermato: in esecuzione installata il database è in `%LOCALAPPDATA%\HDManagerDesktop\data\`; **l’aggiornamento installer sostituisce solo i file in cartella programma**, non cancella i dati utente in LocalAppData.
- Prima installazione: installazione standard; al primo avvio viene creato il DB in LocalAppData se assente.

---

## Note di rilascio v1.0.4

### Interfaccia
- Barra superiore: icona applicazione in formato circolare (senza titolo testuale ripetuto).

### Tab Note Cliente
- **Salva in file e archivio**: salva anche la nota nella **lista note** del cliente (oltre a file e archivio), così i dati restano allineati.

### Installer
- Aggiornamento in-place: in caso di reinstallazione sulla stessa cartella, la schermata della cartella di installazione viene omessa quando possibile (`DisableDirPage=auto`).

---

## Requisiti di sistema

- **Sistema operativo:** Windows 10/11 (64-bit)
- **Spazio disco:** ~150 MB
- **Memoria RAM:** 512 MB minimo

---

## Funzionalità principali

### Gestione entità
- **Competenze** – Classificazione delle competenze aziendali
- **Setup Tipi Prodotti** – Configurazione tipi prodotto con flag (IP, Host, RDP, URL, ecc.)
- **Prodotti** – Catalogo prodotti con associazione clienti e ambienti
- **Ambienti** – Ambienti di deploy (es. Sviluppo, Test, Produzione)
- **Release** – Versioni software
- **Clienti** – Anagrafica clienti con località, link, risorse e VPN associate
- **Risorse** – Persone con ruoli, contatti e competenze
- **Ruoli** – Ruoli aziendali con ordine di visualizzazione
- **VPN** – Connessioni VPN (proprietarie o Windows) con clienti associati

### Interfaccia
- Finestra principale con linguette **Clienti** e **Archivio**
- Vista ad albero **Cliente → Prodotti** con dettaglio contestuale
- Schede cliente: **Info Cliente**, **Accessi**, **Archivio Cliente**, **Rubrica**, **Note**
- **Impostazioni** con menu laterale per tutte le entità
- Inserimento rapido con Ctrl+V da Excel/CSV
- Campi multi-selezione con finestra di scelta valori

### Tab Note Cliente
- **Vista Testo** – Foglio di testo libero
- **Vista Tabella** – Foglio tipo Excel con celle, navigazione con frecce, auto-estensione righe/colonne
- **Formule** – Somma, sottrazione, moltiplicazione, divisione, percentuale, radici, concatenazione, media. Supporto celle sparse e "Applica a tutte le righe"
- **Annulla/Ripristina** – Undo/Redo con Ctrl+Z e Ctrl+Y (fino a 50 operazioni)
- **Anteprima cella** – Riquadro per visualizzare e modificare contenuti lunghi
- **Copia per Excel/Word** – Incolla in Excel (TSV) o Word (tabella HTML)
- **Incolla celle / Incolla come tabella** – Da Excel/Word verso l'app
- **Eliminazione massiva** – Elimina righe/colonne selezionate (anche non consecutive)
- **Ctrl+Shift+frecce** – Selezione fino all'ultima cella compilata

### Accessi e credenziali
- Connessione RDP (IP/Host o file .rdp preconfigurato)
- Gestione credenziali prodotto per ambiente/versione
- VPN Windows integrata (rasdial)
- VPN proprietarie con avvio eseguibile
- Password salvate nel Credential Manager di Windows (keyring)

### Archivio
- Cartelle e sottocartelle
- File e link con tag
- Spostamento file tra cartelle (click destro)
- Filtri per nome, estensione, tag

### Pacchetti
- Export/Import configurazioni in JSON (Core, Risorse, VPN)
- Trasferimento entità tra installazioni

---

## Installazione

1. Scarica `HDManagerDesktop-Setup-1.0.8.exe` dalla release
2. Esegui l'installer (richiesti privilegi amministratore)
3. Segui la procedura guidata
4. Opzionale: crea icona sul desktop

### Aggiornamento
Se HD Manager Desktop è già installato (stesso identificativo applicazione), l'installer **sostituisce solo i file del programma** nella cartella di installazione (tipicamente `Program Files`) e **non** elimina database, log o altri dati in `%LOCALAPPDATA%\HDManagerDesktop\`.

### Nuova installazione
Se il programma non è installato, l'installer propone una **nuova installazione** in una cartella pulita (nessun dato precedente).

---

## Tecnologie

- **Python 3.12**
- **PyQt6** – Interfaccia grafica
- **SQLite** – Database locale
- **keyring** – Storage sicuro password (Windows Credential Manager)
- **openpyxl** – Export note in formato Excel

---

## Note di rilascio v1.0.3

### Tab Note Cliente (nuovo)
- Vista Testo e Vista Tabella con linguette
- Tabella tipo Excel: celle, intestazioni A/B/C e 1/2/3, navigazione con frecce
- Estensione automatica righe/colonne oltre l'ultima cella
- Formule: Somma, Sottrazione, Moltiplicazione, Divisione, Percentuale, Radici (quadrata, cubica, n-esima), Media, Concatenazione righe/colonne
- Applica formula a tutte le righe (come Excel)
- Annulla/Ripristina (Ctrl+Z, Ctrl+Y)
- Anteprima e modifica contenuto cella
- Copia per Excel (TSV) e Word (tabella HTML)
- Incolla celle (punto corrente) e Incolla come tabella
- Elimina righe/colonne selezionate (anche non consecutive)
- Ctrl+Shift+frecce per selezione fino all'ultima cella compilata

### Correzioni e miglioramenti
- Connessioni RDP eseguite in background (v1.0.2)
- PowerShell nascosto all'apertura Impostazioni (v1.0.2)

---

## Note di rilascio v1.0.2

- Connessioni RDP eseguite in background (interfaccia non bloccata)
- PowerShell nascosto all'apertura Impostazioni (lista VPN Windows)
- Corretto salvataggio campo Clienti in Impostazioni > VPN
- Miglioramenti generali di stabilità

---

## Licenza

Progetto privato – HD Manager.
