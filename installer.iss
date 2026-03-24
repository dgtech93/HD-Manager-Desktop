; Inno Setup script - HD Manager Desktop
; Compila con Inno Setup 6.x
;
; Comportamento:
; - Stesso AppId: se HD Manager è già installato, l'installer AGGIORNA in-place (exe e file sotto {app})
;   senza toccare dati utente fuori dalla cartella programma (es. %LOCALAPPDATA%\HDManagerDesktop).
; - Prima installazione: installazione pulita in Program Files (o cartella scelta).
; - Disinstallazione: i dati locali si rimuovono solo se l'utente conferma nella dialog di disinstallazione.
;
; Uso rapido:
; 1) Crea la build dell'app (es. PyInstaller) in: .\dist\HDManagerDesktop\
; 2) Apri questo file in Inno Setup Compiler
; 3) Build -> Compile

#ifndef AppName
  #define AppName "HD Manager Desktop"
#endif
#ifndef AppVersion
  #define AppVersion "1.0.4"
#endif
#ifndef AppPublisher
  #define AppPublisher "HD Manager"
#endif
#ifndef AppExeName
  #define AppExeName "HDManagerDesktop.exe"
#endif
#ifndef BuildRoot
  #define BuildRoot "dist\\HDManagerDesktop"
#endif

[Setup]
AppId={{D1A6B982-77BB-44D1-9B89-8C8E1219D9D0}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\HD Manager Desktop
DefaultGroupName=HD Manager Desktop
DisableDirPage=auto
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
LicenseFile=
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=dist-installer
OutputBaseFilename=HDManagerDesktop-Setup-{#AppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExeName}

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "Crea icona sul desktop"; GroupDescription: "Icone aggiuntive:"; Flags: unchecked

[Files]
; Tutto il contenuto buildato dell'app
Source: "{#BuildRoot}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\HD Manager Desktop"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\HD Manager Desktop"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Avvia HD Manager Desktop"; Flags: nowait postinstall skipifsilent

[Code]
var
  RemoveDataConfirmed: Boolean;

function ConfirmDataRemoval: Boolean;
var
  Msg: string;
begin
  Msg :=
    'Vuoi eliminare anche i dati locali dell''applicazione?' + #13#10 + #13#10 +
    '- Database (entita salvate)' + #13#10 +
    '- Log' + #13#10 +
    '- Cartelle dati correlate all''installazione' + #13#10 + #13#10 +
    'Questa operazione NON e'' reversibile.';

  Result :=
    (MsgBox(Msg, mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES);
end;

procedure DeleteDirIfExists(const DirPath: string);
begin
  if DirExists(DirPath) then
  begin
    DelTree(DirPath, True, True, True);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  InstallRoot: string;
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveDataConfirmed := ConfirmDataRemoval();
  end;

  if (CurUninstallStep = usPostUninstall) and RemoveDataConfirmed then
  begin
    InstallRoot := ExpandConstant('{app}');

    // Dati tipici dentro la cartella installazione
    DeleteDirIfExists(InstallRoot + '\data');
    DeleteDirIfExists(InstallRoot + '\logs');
    DeleteDirIfExists(InstallRoot + '\temp');

    // Eventuali dati utente fuori installazione (se in futuro usati)
    DeleteDirIfExists(ExpandConstant('{localappdata}\HDManagerDesktop'));
    DeleteDirIfExists(ExpandConstant('{userappdata}\HDManagerDesktop'));
    DeleteDirIfExists(ExpandConstant('{commonappdata}\HDManagerDesktop'));
  end;
end;

