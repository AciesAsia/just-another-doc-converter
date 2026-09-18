; DocumentConverter Inno Setup Script
; NZB Consulting
; Version 2.0

#define MyAppName "DocumentConverter"
#define MyAppVersion "2.0"
#define MyAppPublisher "NZB Consulting"
#define MyAppURL "https://github.com/AciesAsia/just-another-doc-converter"
#define MyAppExeName "DocumentConverter.exe"
#define MyAppSourceDir "dist\DocumentConverter"

[Setup]
AppId={AFF4DFC0-E6C2-4E85-A02A-27180FD5619C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=DocumentConverter_Setup_v2.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#MyAppSourceDir}\*";         DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "converter_gui.py";            DestDir: "{app}"; Flags: ignoreversion
Source: "pdf_to_office_gui.py";        DestDir: "{app}"; Flags: ignoreversion
Source: "converter.py";                DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}";           Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";     Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"
Type: files;          Name: "{app}\config.json"

[Code]
var
  OutputFolderPage: TInputDirWizardPage;

procedure InitializeWizard;
begin
  OutputFolderPage := CreateInputDirPage(
    wpSelectDir,
    'Select Output Folder',
    'Where should converted files be saved?',
    'Select the folder where DocumentConverter will save converted files. You can change this later inside the app.',
    False,
    ''
  );
  OutputFolderPage.Add('');
  OutputFolderPage.Values[0] := GetShellFolder('desktop');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath: String;
  OutputFolder: String;
  ConfigLines: TArrayOfString;
begin
  if CurStep = ssPostInstall then
  begin
    ConfigPath := ExpandConstant('{app}\config.json');
    OutputFolder := OutputFolderPage.Values[0];
    StringChange(OutputFolder, '\', '\\');
    SetArrayLength(ConfigLines, 1);
    ConfigLines[0] := '{"output_folder": "' + OutputFolder + '"}';
    SaveStringsToFile(ConfigPath, ConfigLines, False);
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpFinished then
  begin
    WizardForm.FinishedLabel.Caption :=
      'DocumentConverter has been installed.' + #13#10 + #13#10 +
      'Note: Microsoft Office (Word, Excel, PowerPoint) must be installed ' +
      'on this computer for full conversion functionality.';
  end;
end;
