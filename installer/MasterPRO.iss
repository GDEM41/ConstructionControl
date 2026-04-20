#define MyAppName "MasterPRO"
#define MyAppExeName "ConstructionControl.exe"
#define MyPublisher "MasterPRO"
#define SoftMakerPackageDir "SoftMaker.Office.Professional.v2024.1230.1206"
#define PdfPackageDir "PDF-XChange.PRO.v10.8.4.409"

#ifndef MyAppVersion
  #define MyAppVersion "1.0.1.0"
#endif

#ifndef MyAppSourceDir
  #define MyAppSourceDir "..\bin\Release\net10.0-windows"
#endif

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
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=no
SetupLogging=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; Flags: unchecked
Name: "installsoftmaker"; Description: "Установить SoftMaker Office Professional (Portable, рекомендуется для быстрой работы со сметами)"; Flags: unchecked
Name: "installpdfxchange"; Description: "Установить PDF-XChange PRO (рекомендуется для быстрого просмотра и редактирования PDF)"; Flags: unchecked

[Files]
Source: "{#MyAppSourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "packages\{#SoftMakerPackageDir}\*"; DestDir: "{app}\Dependencies\{#SoftMakerPackageDir}"; Flags: recursesubdirs createallsubdirs ignoreversion; Tasks: installsoftmaker
Source: "packages\{#PdfPackageDir}\*"; DestDir: "{tmp}\{#PdfPackageDir}"; Flags: recursesubdirs createallsubdirs deleteafterinstall ignoreversion; Tasks: installpdfxchange

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
var
  DependencyInstallPage: TOutputProgressWizardPage;
  DependencySummary: string;

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

function DetectSoftMakerExecutablePath(): string;
var
  Candidates: array[0..7] of string;
  I: Integer;
begin
  Candidates[0] := ExpandConstant('{app}\Dependencies\{#SoftMakerPackageDir}\PlanMaker.exe');
  Candidates[1] := ExpandConstant('{app}\Dependencies\{#SoftMakerPackageDir}\program\PlanMaker.exe');
  Candidates[2] := ExpandConstant('{pf}\SoftMaker Office Professional 2024\PlanMaker.exe');
  Candidates[3] := ExpandConstant('{pf}\SoftMaker Office Professional 2024\program\PlanMaker.exe');
  Candidates[4] := ExpandConstant('{pf32}\SoftMaker Office Professional 2024\PlanMaker.exe');
  Candidates[5] := ExpandConstant('{pf32}\SoftMaker Office Professional 2024\program\PlanMaker.exe');
  Candidates[6] := ExpandConstant('{pf}\SoftMaker FreeOffice 2024\PlanMaker.exe');
  Candidates[7] := ExpandConstant('{pf32}\SoftMaker FreeOffice 2024\PlanMaker.exe');

  Result := '';
  for I := 0 to GetArrayLength(Candidates) - 1 do
  begin
    if FileExists(Candidates[I]) then
    begin
      Result := Candidates[I];
      Exit;
    end;
  end;
end;

function DetectPdfXChangeExecutablePath(): string;
var
  Candidates: array[0..5] of string;
  I: Integer;
begin
  Candidates[0] := ExpandConstant('{pf}\Tracker Software\PDF Editor\PDFXEdit.exe');
  Candidates[1] := ExpandConstant('{pf}\Tracker Software\PDF Editor\PDFXEdit64.exe');
  Candidates[2] := ExpandConstant('{pf32}\Tracker Software\PDF Editor\PDFXEdit.exe');
  Candidates[3] := ExpandConstant('{pf}\Tracker Software\PDF-XChange Editor\PDFXEdit.exe');
  Candidates[4] := ExpandConstant('{pf}\Tracker Software\PDF-XChange Editor\PDFXEdit64.exe');
  Candidates[5] := ExpandConstant('{pf32}\Tracker Software\PDF-XChange Editor\PDFXEdit.exe');

  Result := '';
  for I := 0 to GetArrayLength(Candidates) - 1 do
  begin
    if FileExists(Candidates[I]) then
    begin
      Result := Candidates[I];
      Exit;
    end;
  end;
end;

function IsSoftMakerInstalled(): Boolean;
begin
  Result :=
    IsInstalledByDisplayName('SoftMaker') or
    IsInstalledByDisplayName('PlanMaker') or
    (DetectSoftMakerExecutablePath() <> '');
end;

function IsPdfXChangeInstalled(): Boolean;
begin
  Result :=
    IsInstalledByDisplayName('PDF-XChange') or
    (DetectPdfXChangeExecutablePath() <> '');
end;

procedure AppendDependencySummary(const Line: string);
begin
  if DependencySummary <> '' then
    DependencySummary := DependencySummary + #13#10;

  DependencySummary := DependencySummary + Line;
end;

function FindTaskIndexByText(const NamePart: string): Integer;
var
  I: Integer;
begin
  Result := -1;
  for I := 0 to WizardForm.TasksList.Items.Count - 1 do
  begin
    if Pos(Uppercase(NamePart), Uppercase(WizardForm.TasksList.Items[I])) > 0 then
    begin
      Result := I;
      Exit;
    end;
  end;
end;

procedure MarkDependencyProgress(const Title, Status: string; Position, Total: Integer);
begin
  DependencyInstallPage.SetText(Title, Status);
  DependencyInstallPage.SetProgress(Position, Total);
end;

function RunDependencyInstaller(const DisplayName, CommandPath, WorkingDir: string; StepIndex, TotalSteps: Integer): Boolean;
var
  ResultCode: Integer;
begin
  Result := False;

  if not FileExists(CommandPath) then
  begin
    AppendDependencySummary(DisplayName + ': файл установки не найден.');
    Exit;
  end;

  MarkDependencyProgress(DisplayName, 'Выполняется тихая установка...', StepIndex - 1, TotalSteps);

  if not Exec(ExpandConstant('{cmd}'), '/C ' + AddQuotes(CommandPath), WorkingDir, SW_SHOW, ewWaitUntilTerminated, ResultCode) then
  begin
    AppendDependencySummary(DisplayName + ': не удалось запустить установку.');
    Exit;
  end;

  if ResultCode <> 0 then
  begin
    AppendDependencySummary(DisplayName + ': установка завершилась с кодом ' + IntToStr(ResultCode) + '.');
    Exit;
  end;

  MarkDependencyProgress(DisplayName, 'Установлено успешно.', StepIndex, TotalSteps);
  AppendDependencySummary(DisplayName + ': установлено успешно.');
  Result := True;
end;

procedure InstallSelectedDependencies();
var
  TotalSteps: Integer;
  StepIndex: Integer;
  SoftMakerDir: string;
  PdfDir: string;
begin
  TotalSteps := 0;
  if WizardIsTaskSelected('installsoftmaker') then
    TotalSteps := TotalSteps + 1;
  if WizardIsTaskSelected('installpdfxchange') then
    TotalSteps := TotalSteps + 1;

  if TotalSteps = 0 then
    Exit;

  StepIndex := 1;
  DependencySummary := '';
  DependencyInstallPage.Show;
  try
    if WizardIsTaskSelected('installsoftmaker') then
    begin
      SoftMakerDir := ExpandConstant('{app}\Dependencies\{#SoftMakerPackageDir}');
      RunDependencyInstaller(
        'SoftMaker Office Professional',
        AddBackslash(SoftMakerDir) + 'PORTABLE.cmd',
        SoftMakerDir,
        StepIndex,
        TotalSteps);
      StepIndex := StepIndex + 1;
    end;

    if WizardIsTaskSelected('installpdfxchange') then
    begin
      PdfDir := ExpandConstant('{tmp}\{#PdfPackageDir}');
      RunDependencyInstaller(
        'PDF-XChange PRO',
        AddBackslash(PdfDir) + 'INSTALL.cmd',
        PdfDir,
        StepIndex,
        TotalSteps);
      StepIndex := StepIndex + 1;
    end;

    MarkDependencyProgress('Дополнительные программы', 'Установка завершена.', TotalSteps, TotalSteps);
  finally
    DependencyInstallPage.Hide;
  end;

  if DependencySummary <> '' then
    MsgBox(DependencySummary, mbInformation, MB_OK);
end;

procedure InitializeWizard();
var
  SoftMakerTaskIndex: Integer;
  PdfTaskIndex: Integer;
begin
  DependencyInstallPage := CreateOutputProgressPage(
    'Установка дополнительных программ',
    'Подождите, пока будут установлены выбранные компоненты.');

  WizardForm.WelcomeLabel1.Visible := False;
  WizardForm.WelcomeLabel2.Visible := False;
  WizardForm.FinishedHeadingLabel.Visible := False;
  WizardForm.FinishedLabel.Visible := False;
  WizardForm.WelcomeLabel1.Caption := '';
  WizardForm.WelcomeLabel2.Caption := '';
  WizardForm.FinishedHeadingLabel.Caption := '';
  WizardForm.FinishedLabel.Caption := '';

  SoftMakerTaskIndex := FindTaskIndexByText('SoftMaker Office Professional');
  PdfTaskIndex := FindTaskIndexByText('PDF-XChange PRO');

  if SoftMakerTaskIndex >= 0 then
    WizardForm.TasksList.Checked[SoftMakerTaskIndex] := not IsSoftMakerInstalled;

  if PdfTaskIndex >= 0 then
    WizardForm.TasksList.Checked[PdfTaskIndex] := not IsPdfXChangeInstalled;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    InstallSelectedDependencies();
end;
