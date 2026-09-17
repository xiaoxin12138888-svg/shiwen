[Setup]
AppId={{AF647C90-3404-4A08-88F5-5C1F4538796B}
AppName=拾文
AppVersion=1.0.0
AppPublisher=拾文个人工具
DefaultDirName={localappdata}\Programs\Shiwen
DefaultGroupName=拾文
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
OutputDir=..\release
OutputBaseFilename=拾文-Setup-1.0.0-Windows-x64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\Shiwen.exe
CloseApplications=yes
RestartApplications=no
SetupLogging=yes
DisableProgramGroupPage=yes
InfoAfterFile=使用说明.txt

[Languages]
Name: "chinesesimplified"; MessagesFile: "..\build\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："

[Files]
Source: "..\build\bundle\Shiwen\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "使用说明.txt"; DestDir: "{app}"; DestName: "使用说明.txt"; Flags: ignoreversion

[Icons]
Name: "{group}\拾文"; Filename: "{app}\Shiwen.exe"
Name: "{group}\退出拾文"; Filename: "{app}\Shiwen.exe"; Parameters: "--shutdown"
Name: "{group}\使用说明"; Filename: "{app}\使用说明.txt"
Name: "{group}\卸载拾文"; Filename: "{uninstallexe}"
Name: "{autodesktop}\拾文"; Filename: "{app}\Shiwen.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Shiwen.exe"; Description: "立即打开拾文"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{app}\Shiwen.exe"; Parameters: "--shutdown"; Flags: runhidden waituntilterminated; RunOnceId: "StopShiwen"

[Code]
function PrepareToInstall(var NeedsRestart: Boolean): String;
var Code: Integer;
begin
  Result := '';
  if FileExists(ExpandConstant('{app}\Shiwen.exe')) then
    Exec(ExpandConstant('{app}\Shiwen.exe'), '--shutdown', '', SW_HIDE, ewWaitUntilTerminated, Code);
end;
