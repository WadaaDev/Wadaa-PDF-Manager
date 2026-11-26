#define AppName "Wadaa PDF Manager"
#define VersionHandle FileOpen(SourcePath + "VERSION.txt")
#define AppVersion Trim(FileRead(VersionHandle))
#expr FileClose(VersionHandle)
#if AppVersion == ""
  #error VERSION.txt is empty.
#endif
#define BuiltVersionHandle FileOpen(SourcePath + "dist\" + AppName + "\_internal\VERSION.txt")
#define BuiltVersion Trim(FileRead(BuiltVersionHandle))
#expr FileClose(BuiltVersionHandle)
#if BuiltVersion != AppVersion
  #error Rebuild the application before compiling Setup: VERSION.txt and the bundled version differ.
#endif
#define AppExe AppName + ".exe"

[Setup]
AppId={{C59DF9FC-A134-448E-ADEC-43C04E72FA8E}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={autopf}\{#AppName}
DisableDirPage=no
UsePreviousAppDir=yes
AppendDefaultDirName=no
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#AppExe}
SetupIconFile=assets\icon.ico
OutputDir=installer
OutputBaseFilename=Wadaa PDF Manager Setup v{#AppVersion}
WizardStyle=modern
VersionInfoVersion={#AppVersion}.0
VersionInfoProductVersion={#AppVersion}.0
VersionInfoProductName={#AppName}
VersionInfoCompany=WadaaDev

[Files]
Source: "dist\{#AppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"; IconFilename: "{app}\{#AppExe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName}"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent runasoriginaluser; Check: ShouldLaunchApp

[Code]
const
  AppFolderName = '{#AppName}';

var
  LastDisplayedDir: String;

function ShouldLaunchApp: Boolean;
begin
  Result := True;
end;

function CleanInstallPath(Path: String): String;
begin
  Result := RemoveBackslashUnlessRoot(PathNormalizeSlashes(Trim(Path)));
  if (Length(Result) = 2) and (Result[2] = ':') then
    Result := AddBackslash(Result);
end;

function IsInstallRoot(Path: String): Boolean;
var
  Drive: String;
begin
  Drive := ExtractFileDrive(Path);
  Result := (Drive <> '') and
    ((CompareText(Path, Drive) = 0) or
     (CompareText(Path, AddBackslash(Drive)) = 0));
end;

function HasAppFolderName(Path: String): Boolean;
begin
  Result := (not IsInstallRoot(Path)) and
    (CompareText(ExtractFileName(Path), AppFolderName) = 0);
end;

function InstallPathFromParent(Parent: String): String;
begin
  Parent := CleanInstallPath(Parent);
  while HasAppFolderName(Parent) do
    Parent := CleanInstallPath(ExtractFileDir(Parent));
  Result := AddBackslash(Parent) + AppFolderName;
end;

function CorrectTypedInstallPath(Path, DisplayedDir, PreviousDir: String): String;
var
  Parent: String;
  DisplayedParent: String;
begin
  Result := Path;
  if (PreviousDir <> '') and (CompareText(Path, PreviousDir) = 0) then
    Exit;

  if HasAppFolderName(Path) or IsInstallRoot(Path) then
    Result := InstallPathFromParent(Path)
  else
  begin
    Parent := CleanInstallPath(ExtractFileDir(Path));
    DisplayedParent := CleanInstallPath(ExtractFileDir(DisplayedDir));

    if (CompareText(Path, DisplayedParent) = 0) or
       (IsInstallRoot(Parent) and
        (CompareText(Parent, DisplayedParent) <> 0)) then
      Result := InstallPathFromParent(Path)
    else
      Result := InstallPathFromParent(Parent);
  end;
end;

procedure BrowseInstallParent(Sender: TObject);
var
  Parent: String;
begin
  Parent := CleanInstallPath(WizardForm.DirEdit.Text);

  if HasAppFolderName(Parent) or
     (CompareText(Parent, CleanInstallPath(WizardForm.PrevAppDir)) = 0) then
    Parent := CleanInstallPath(ExtractFileDir(Parent));

  if BrowseForFolder('Choose the parent folder for ' + AppFolderName, Parent, False) then
  begin
    LastDisplayedDir := InstallPathFromParent(Parent);
    WizardForm.DirEdit.Text := LastDisplayedDir;
  end;
end;

procedure InitializeWizard;
begin
  LastDisplayedDir := CleanInstallPath(WizardForm.DirEdit.Text);
  WizardForm.DirBrowseButton.OnClick := @BrowseInstallParent;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  SelectedPath: String;
begin
  Result := True;

  if CurPageID = wpSelectDir then
  begin
    SelectedPath := CleanInstallPath(WizardForm.DirEdit.Text);

    if ExtractFileDrive(SelectedPath) = '' then
    begin
      SuppressibleMsgBox('Choose an installation folder.', mbError, MB_OK, IDOK);
      Result := False;
      Exit;
    end;

    LastDisplayedDir := CorrectTypedInstallPath(
      SelectedPath,
      LastDisplayedDir,
      CleanInstallPath(WizardForm.PrevAppDir)
    );

    WizardForm.DirEdit.Text := LastDisplayedDir;
  end;
end;