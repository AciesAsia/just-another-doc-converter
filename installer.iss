; installer.iss — Inno Setup script for DocumentConverter
; Produces: DocumentConverter_Setup_v2.0.exe
;
; Requirements:
;   - Inno Setup 6.x  (https://jrsoftware.org/isdl.php)
;   - PyInstaller output must exist at:  dist\DocumentConverter\
;
; To compile:
;   Open this file in Inno Setup Compiler and click Build > Compile
;   OR from command line:
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss

#define MyAppName      "DocumentConverter"
#define MyAppVersion   "2.0"
#define MyAppPublisher "NZB Consulting"
#define MyAppURL       "https://github.com/YOUR-USERNAME/DocumentConverter"
#define MyAppExeName   "DocumentConverter.exe"
#define MyAppSourceDir "dist\DocumentConverter"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} v{#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=DocumentConverter_Setup_v{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=admin
MinVersion=10.0
WizardStyle=modern
WizardSizePercent=120
; F-29: close any running instance before install
CloseApplications=yes
; F-12: show GPL-3.0 licence during installation
LicenseFile=LICENSE
; F-13: desktop shortcut defaults to unchecked — set below in [Tasks]

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; F-13: desktop shortcut is opt-in (unchecked by default)
Name: "desktopicon"; Description: "Create a &desktop shortcut"; \
    GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "{#MyAppSourceDir}\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; F-11: changed DestPath: (invalid key) to Filename: (correct Inno Setup key)
Name: "{group}\{#MyAppName}";           Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";     Filename: "{app}\{#MyAppExeName}"; \
    Tasks: desktopicon

[Code]
{ ── Custom output folder page ────────────────────────────────────────────── }
var
  OutputFolderPage: TInputDirWizardPage;

procedure InitializeWizard;
begin
  { Create a custom directory-selection page after the Tasks page }
  OutputFolderPage := CreateInputDirPage(
    wpSelectTasks,
    'Select Output Folder',
    'Where should converted files be saved?',
    'DocumentConverter will save all converted files to this folder.'
    + #13#10
    + 'You can change this at any time using the Browse button inside the app.',
    False,
    ''
  );
  OutputFolderPage.Add('Output folder:');

  { Default to the user''s Desktop }
  OutputFolderPage.Values[0] := ExpandConstant('{userdesktop}');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath, JsonContent, OutputFolder: String;
  InputDir, LogDir: String;
begin
  { ── F-28: ssInstall — create app subdirectories during file installation ── }
  if CurStep = ssInstall then
  begin
    InputDir := ExpandConstant('{app}\input');
    LogDir   := ExpandConstant('{app}\logs');
    if not DirExists(InputDir) then
      CreateDir(InputDir);
    if not DirExists(LogDir) then
      CreateDir(LogDir);
  end;

  { ── ssPostInstall — write config.json with chosen output folder ─────────── }
  if CurStep = ssPostInstall then
  begin
    OutputFolder := OutputFolderPage.Values[0];
    { Escape backslashes for JSON }
    StringChangeEx(OutputFolder, '\', '\\', True);

    JsonContent := '{'                                    + #13#10
                 + '  "output_folder": "' + OutputFolder + '",' + #13#10
                 + '  "save_log": true'                          + #13#10
                 + '}';

    ConfigPath := ExpandConstant('{app}\config.json');
    SaveStringToFile(ConfigPath, JsonContent, False);
  end;
end;

{ ── Run section ─────────────────────────────────────────────────────────── }
[Run]
Filename: "{app}\{#MyAppExeName}"; \
    Description: "Launch {#MyAppName} now"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; F-38: clean up runtime-generated files and directories on uninstall
Type: filesandordirs; Name: "{app}\logs"
Type: filesandordirs; Name: "{app}\input"
Type: files;          Name: "{app}\config.json"

[Messages]
FinishedLabel=DocumentConverter has been installed.%n%nNote: Microsoft Office (Word, Excel, PowerPoint) must be installed on this machine for all conversion features to work.
