#define MyAppName "МастерPRO"
#define MyAppExeName "ConstructionControl.exe"
#define MyAppVersion "1.0.0.0"
#define MyPublisher "МастерPRO"

[Setup]
AppId={{2B6C2E46-0D7A-4B08-9B1E-12E0D6F74CF1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=output
OutputBaseFilename=MasterPRO_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\MasterPRO.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
WizardImageFile=assets\wizard.bmp
WizardSmallImageFile=assets\wizard_small.bmp
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=no
SetupLogging=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; Flags: unchecked
Name: "installpdfxchange"; Description: "Открыть страницу установки PDF‑XChange Editor"; Flags: unchecked
Name: "installplanmaker"; Description: "Открыть страницу установки PlanMaker (FreeOffice)"; Flags: unchecked

[Files]
Source: "..\ConstructionControl\bin\Release\net10.0-windows\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "https://www.pdf-xchange.com/product/pdf-xchange-editor"; Description: "Открыть страницу PDF‑XChange Editor"; Flags: shellexec postinstall nowait; Tasks: installpdfxchange
Filename: "https://www.softmaker.com/en/freeoffice"; Description: "Открыть страницу PlanMaker (FreeOffice)"; Flags: shellexec postinstall nowait; Tasks: installplanmaker

[Code]
function IsInstalledInRoot(RootKey: Integer; const NamePart: string): Boolean;
var
  UninstKey: string;
  SubKeys: TArrayOfString;
  I: Integer;
  DispName: string;
begin
  Result := False;
  UninstKey := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';
  if not RegGetSubkeyNames(RootKey, UninstKey, SubKeys) then
    Exit;

  for I := 0 to GetArrayLength(SubKeys) - 1 do
  begin
    if RegQueryStringValue(RootKey, UninstKey + '\' + SubKeys[I], 'DisplayName', DispName) then
    begin
      if Pos(Uppercase(NamePart), Uppercase(DispName)) > 0 then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

function IsInstalledByDisplayName(const NamePart: string): Boolean;
begin
  Result :=
    IsInstalledInRoot(HKLM, NamePart) or
    IsInstalledInRoot(HKLM64, NamePart) or
    IsInstalledInRoot(HKCU, NamePart);
end;

function IsPdfXChangeInstalled: Boolean;
begin
  Result := IsInstalledByDisplayName('PDF-XChange');
end;

function IsPlanMakerInstalled: Boolean;
begin
  Result := IsInstalledByDisplayName('PlanMaker');
end;

procedure InitializeWizard();
var
  PdfTask: Integer;
  PlanTask: Integer;
begin
  PdfTask := WizardForm.TasksList.Items.IndexOf('Открыть страницу установки PDF‑XChange Editor');
  PlanTask := WizardForm.TasksList.Items.IndexOf('Открыть страницу установки PlanMaker (FreeOffice)');

  if PdfTask >= 0 then
    WizardForm.TasksList.Checked[PdfTask] := not IsPdfXChangeInstalled;

  if PlanTask >= 0 then
    WizardForm.TasksList.Checked[PlanTask] := not IsPlanMakerInstalled;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if (not IsPdfXChangeInstalled) or (not IsPlanMakerInstalled) then
      MsgBox('Рекомендуется установить PDF‑XChange Editor и PlanMaker для полноценной работы МастерPRO.', mbInformation, MB_OK);
  end;
end;
