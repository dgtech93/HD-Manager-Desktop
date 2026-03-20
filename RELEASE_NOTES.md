# HD Manager Desktop - Release v1.0.2

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
- Schede cliente: **Info Cliente**, **Accessi**, **Archivio Cliente**
- **Impostazioni** con menu laterale per tutte le entità
- Inserimento rapido con Ctrl+V da Excel/CSV
- Campi multi-selezione con finestra di scelta valori

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

1. Scarica `HDManagerDesktop-Setup.exe` dalla release
2. Esegui l’installer (richiesti privilegi amministratore)
3. Segui la procedura guidata
4. Opzionale: crea icona sul desktop

### Aggiornamento
Se HD Manager Desktop è già installato, l’installer aggiorna i file mantenendo i dati esistenti (database e log in `%LOCALAPPDATA%\HDManagerDesktop\`).

---

## Tecnologie

- **Python 3.12**
- **PyQt6** – Interfaccia grafica
- **SQLite** – Database locale
- **keyring** – Storage sicuro password (Windows Credential Manager)

---

## Note di rilascio v1.0.2

- Connessioni RDP eseguite in background (interfaccia non bloccata)
- PowerShell nascosto all’apertura Impostazioni (lista VPN Windows)
- Corretto salvataggio campo Clienti in Impostazioni > VPN
- Miglioramenti generali di stabilità

---

## Licenza

Progetto privato – HD Manager.
