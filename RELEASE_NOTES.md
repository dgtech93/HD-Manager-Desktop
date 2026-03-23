# HD Manager Desktop - Release v1.0.3

Applicazione desktop per la gestione di clienti, prodotti, risorse, VPN e archivio documentale.

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

1. Scarica `HDManagerDesktop-Setup-1.0.3.exe` dalla release
2. Esegui l'installer (richiesti privilegi amministratore)
3. Segui la procedura guidata
4. Opzionale: crea icona sul desktop

### Aggiornamento
Se HD Manager Desktop è già installato, l'installer **aggiorna i file** mantenendo i dati esistenti (database e log in `%LOCALAPPDATA%\HDManagerDesktop\`). Nessuna perdita di dati.

### Nuova installazione
Se il programma non è installato, l'installer esegue una **installazione pulita**.

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
