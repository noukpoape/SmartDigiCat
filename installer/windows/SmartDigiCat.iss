; ---------------------------------------------------------------------------
; Installeur Windows de SmartDigiCat (Inno Setup 6)
;
; Produit un seul fichier : SmartDigiCat-Windows-Setup.exe (nom fixe : le lien
; « releases/latest/download/... » du README reste valable a chaque version)
; - installation hors ligne (Python et bibliotheques sont dans l'installeur) ;
; - pas de droits administrateur necessaires (installation pour l'utilisateur),
;   ou pour tous les utilisateurs si on le choisit au lancement ;
; - assistant classique : Bienvenue > Informations > Dossier > Menu Demarrer >
;   Taches (icone du Bureau) > Pret a installer > Installation > Termine ;
; - raccourcis : menu Demarrer (+ Desinstaller) et, au choix, bureau ;
; - dossier d'installation et fichiers MASQUES ;
; - desinstallation : Parametres > Applications, ou menu Demarrer.
;
; Compilation (depuis la racine du depot, apres PyInstaller) :
;   iscc /DMyAppVersion=1.0.0 installer\windows\SmartDigiCat.iss
; ---------------------------------------------------------------------------

#define MyAppName "SmartDigiCat"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "Centrale Lille"
#define MyAppURL "https://github.com/noukpoape/SmartDigiCat"
#define MyAppExeName "SmartDigiCat.exe"

[Setup]
AppId={{6F0B3C2A-5D7E-4C1B-9A2F-3E8D1B7C4A90}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
; --- etapes de l'assistant (toutes affichees)
DisableWelcomePage=no
DisableDirPage=no
DisableProgramGroupPage=no
AllowNoIcons=yes
DisableReadyPage=no
DisableFinishedPage=no
; Licence : quand le fichier LICENSE existera, ajouter une page d'acceptation avec
; LicenseFile=..\..\LICENSE
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\..\dist
OutputBaseFilename={#MyAppName}-Windows-Setup
SetupIconFile=..\icons\smartdigicat.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"; InfoBeforeFile: "info_fr.txt"
Name: "english"; MessagesFile: "compiler:Default.isl"; InfoBeforeFile: "info_en.txt"

[CustomMessages]
french.DossierNonVide=Ce dossier contient déjà d'autres fichiers.%n%nChoisissez un dossier vide ou un nouveau dossier pour SmartDigiCat (par exemple, ajoutez \SmartDigiCat à la fin du chemin).
english.DossierNonVide=This folder already contains other files.%n%nChoose an empty or new folder for SmartDigiCat (for example, add \SmartDigiCat at the end of the path).

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Dirs]
; dossier d'installation masque
Name: "{app}"; Attribs: hidden

[Files]
; tout le programme construit par PyInstaller, fichiers masques
Source: "..\..\dist\{#MyAppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Attribs: hidden

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; fichiers crees apres l'installation (caches Python...) : uniquement ceux du logiciel
Type: filesandordirs; Name: "{app}\_internal"
Type: dirifempty; Name: "{app}"

[Code]
{ Le dossier d'installation est masque et vide a la desinstallation : il doit donc etre
  vide, nouveau, ou contenir deja SmartDigiCat (mise a jour). On refuse par exemple
  « Documents », qui serait sinon masque. }
function DossierVide(const Dossier: String): Boolean;
var
  Recherche: TFindRec;
begin
  Result := True;
  if FindFirst(AddBackslash(Dossier) + '*', Recherche) then
  begin
    try
      repeat
        if (Recherche.Name <> '.') and (Recherche.Name <> '..') then
        begin
          Result := False;
          Exit;
        end;
      until not FindNext(Recherche);
    finally
      FindClose(Recherche);
    end;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Dossier: String;
begin
  Result := True;
  if CurPageID = wpSelectDir then
  begin
    Dossier := WizardDirValue;
    if DirExists(Dossier) and not FileExists(AddBackslash(Dossier) + '{#MyAppExeName}') and not DossierVide(Dossier) then
    begin
      MsgBox(CustomMessage('DossierNonVide'), mbError, MB_OK);
      Result := False;
    end;
  end;
end;
