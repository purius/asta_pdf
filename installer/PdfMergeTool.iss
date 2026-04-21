#define AppName "PDF 뷰어"
#define AppPublisher "PdfMergeTool"
#define AppExeName "PdfMergeTool.exe"
#define ProductId "PdfMergeTool"
#define PdfProgId "PdfMergeTool.Pdf"

[Setup]
AppId={{839FC012-54EC-43E9-9BB9-AD6E3C26A428}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
UninstallDisplayName={#AppName}
DefaultDirName={localappdata}\Programs\{#ProductId}
DefaultGroupName={#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename=PdfMergeToolSetup
SetupIconFile={#RootDir}\src\PdfMergeTool\Assets\AppIcon.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no
AppMutex=PdfMergeTool.SingleInstance
ChangesAssociations=yes
SetupLogging=yes
ShowLanguageDialog=no
VersionInfoVersion={#AppVersion}.0
VersionInfoCompany={#AppPublisher}
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoDescription={#AppName} 설치

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕 화면 바로가기 만들기"; GroupDescription: "추가 작업:"; Flags: checkedonce

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
Type: filesandordirs; Name: "{app}\*"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Uninstall\{#ProductId}"; Flags: deletekey

Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\{#ProductId}"; ValueType: string; ValueName: ""; ValueData: "PDF 통합..."; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\{#ProductId}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\{#ProductId}"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\{#ProductId}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" --merge ""%1"""; Flags: uninsdeletekey

Root: HKCU; Subkey: "Software\Classes\.pdf\shell\{#ProductId}"; ValueType: string; ValueName: ""; ValueData: "PDF 통합..."; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\.pdf\shell\{#ProductId}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\.pdf\shell\{#ProductId}"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\.pdf\shell\{#ProductId}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" --merge ""%1"""; Flags: uninsdeletekey

Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}\shell\{#ProductId}"; ValueType: string; ValueName: ""; ValueData: "PDF 통합..."; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}\shell\{#ProductId}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}\shell\{#ProductId}"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}\shell\{#ProductId}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" --merge ""%1"""; Flags: uninsdeletekey

Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}"; ValueType: string; ValueName: ""; ValueData: "PDF 문서"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}"; ValueType: string; ValueName: "FriendlyTypeName"; ValueData: "PDF 문서"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}"; ValueType: dword; ValueName: "EditFlags"; ValueData: "$00000000"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"",0"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#PdfProgId}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Flags: uninsdeletekey

Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"",0"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; ValueName: ".pdf"; ValueData: ""; Flags: uninsdeletekey

Root: HKCU; Subkey: "Software\Classes\.pdf"; ValueType: string; ValueName: ""; ValueData: "{#PdfProgId}"
Root: HKCU; Subkey: "Software\Classes\.pdf"; ValueType: string; ValueName: "Content Type"; ValueData: "application/pdf"
Root: HKCU; Subkey: "Software\Classes\.pdf"; ValueType: string; ValueName: "PerceivedType"; ValueData: "document"
Root: HKCU; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: binary; ValueName: "{#PdfProgId}"; ValueData: ""
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithProgids"; ValueType: binary; ValueName: "{#PdfProgId}"; ValueData: ""

Root: HKCU; Subkey: "Software\{#ProductId}\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\{#ProductId}\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "PDF 보기, 페이지 정리, 통합 도구"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\{#ProductId}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pdf"; ValueData: "{#PdfProgId}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "{#ProductId}"; ValueData: "Software\{#ProductId}\Capabilities"; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[Code]
const
  PdfProgId = '{#PdfProgId}';
  ProductId = '{#ProductId}';

procedure StopRunningApp;
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{cmd}'), '/c taskkill /IM {#AppExeName} /F /T >nul 2>nul', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function InitializeSetup: Boolean;
begin
  StopRunningApp;
  Result := True;
end;

function InitializeUninstall: Boolean;
begin
  StopRunningApp;
  Result := True;
end;

procedure RemovePdfDefaultIfOwned;
var
  CurrentProgId: string;
begin
  if RegQueryStringValue(HKCU, 'Software\Classes\.pdf', '', CurrentProgId) then
  begin
    if CompareText(CurrentProgId, PdfProgId) = 0 then
      RegDeleteValue(HKCU, 'Software\Classes\.pdf', '');
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RemovePdfDefaultIfOwned;
    RegDeleteValue(HKCU, 'Software\Classes\.pdf\OpenWithProgids', PdfProgId);
    RegDeleteValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithProgids', PdfProgId);
    RegDeleteValue(HKCU, 'Software\RegisteredApplications', ProductId);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\SystemFileAssociations\.pdf\shell\' + ProductId);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\.pdf\shell\' + ProductId);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\' + PdfProgId);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Applications\{#AppExeName}');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\' + ProductId);
  end;
end;
