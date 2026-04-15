# APPDesktopHDManager

**Versione corrente:** v1.0.7 (vedi `app/version.py`, `RELEASE_NOTES.md`).

Applicazione desktop Python (PyQt6 + SQLite) per gestire:
- Competenze
- Setup Tipi Prodotti
- Prodotti
- Ambienti
- Release
- Clienti
- Risorse
- Ruoli
- VPN

## UX principale

- Finestra principale con bottone `Impostazioni` in alto a sinistra
- Linguette principali: `Clienti` e `Archivio`
- Nella linguetta `Clienti` viene mostrata la lista clienti e il dettaglio del cliente selezionato
- Nel dettaglio cliente sono presenti le sotto-linguette: `Info Cliente`, `Accessi`, `Archivio Cliente`
- In `Info Cliente` sono visibili `Nome`, `Localita`, risorse divise per ruolo e area prodotti collegati
- Nella linguetta `Archivio` e presente un side menu con voci `Preferiti`, `TAG`, `Archivio`
- In `Impostazioni > Clienti` puoi associare piu risorse al cliente con selezione multipla in tabella
- In `Impostazioni > Ruoli` e disponibile `Ordine visualizzazione` (1-20, univoco) usato per ordinare le sezioni ruolo in `Info Cliente`
- In `Impostazioni > Prodotti` puoi associare uno o piu clienti al prodotto
- In `Clienti` la lista e ad albero (`Cliente -> Prodotti`) e la scheda `Accessi` aggiorna VPN e credenziali in base a cliente/prodotto selezionato
- Con click destro su un prodotto nell'albero clienti puoi creare una nuova credenziale prodotto
- Apertura di una finestra dedicata `Impostazioni` con side menu
- Inserimento/modifica direttamente in tabella
- I campi collegati usano combobox direttamente nella cella
- Per eventuali campi multi-selezione: doppio click sulla cella apre una finestrella di selezione valori
- I campi booleani `0/1` sono mostrati come flag checkbox
- `Nuova riga` aggiunge direttamente una riga pronta alla compilazione
- Dopo `Nuova riga`, con freccia giu sull'ultima riga viene creata automaticamente la riga successiva (inserimento rapido)
- ID numerici non visibili: ogni record mostra un codice alfanumerico univoco
- La modifica righe e disponibile con bottone `Modifica` attivo
- `Nuova riga` abilita automaticamente la modalita modifica per l'inserimento
- Presente bottone `Elimina` su ogni tabella

## Regole dati

- Competenze: `Nome` obbligatorio
- Tipi Prodotti: `Nome` obbligatorio; i flag non impostati vanno a `false`
- Prodotti: tutti i campi obbligatori (`Nome`, `Tipo Prodotto`)
- Ambienti: `Nome` obbligatorio
- Release: `Nome` obbligatorio
- Clienti: `Nome` obbligatorio; `Localita` default `Italia` se vuota
- Risorse: obbligatori `Nome`, `Cognome`, `Ruolo`
- Ruoli: `Nome` obbligatorio; `Piu Clienti` default `false`
- VPN: obbligatori `Nome Connessione`, `Indirizzo Server`, `Tipo VPN`, `Nome Utente`, `Password`
- Tipo VPN consentiti: `Vpn Proprietario`, `VPN Windows`

## Avvio

```bash
pip install -r requirements.txt
python main.py
```

## Test inserimento DB

```bash
python tests/smoke_insert.py
```

Il database SQLite viene creato in `data/app.db`.

## Sicurezza password (VPN / Credenziali)

Se disponibile, l'app salva le password nel **Credential Manager di Windows** tramite la libreria `keyring`,
e nel DB mantiene un campo `password_ref` (fallback: password in chiaro se keyring non disponibile).

## Build installer Windows (Inno Setup)

Prerequisiti:
- Inno Setup 6 installato (`ISCC.exe`)
- PyInstaller installato (`python -m pip install pyinstaller`)

Comando build completo (exe + installer):

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_release.ps1 -Version 1.0.7
```

Output:
- Build app: `dist\HDManagerDesktop\`
- Installer: `dist-installer\HDManagerDesktop-Setup-1.0.7.exe` (il numero versione segue il parametro `-Version`)

### Installazione e aggiornamento

- **Nuova installazione:** eseguire l’installer; il database viene creato in `%LOCALAPPDATA%\HDManagerDesktop\data\` (non in Program Files).
- **Aggiornamento:** stesso `AppId` Inno Setup → installazione in-place su `{app}`; **i dati in LocalAppData non vengono rimossi** dall’aggiornamento.
- **Disinstallazione:** rimozione opzionale dei dati locali solo se confermata dall’utente (vedi `installer.iss`).
