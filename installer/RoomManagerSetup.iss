#define MyAppName "RoomManager for Revit 2026"
#define MyAppVersion "0.7.2"
#define MyAppPublisher "RoomManager Team"
#define MyAppURL "https://example.local/roommanager"
#ifndef PayloadDir
  #define PayloadDir "..\\release\\RoomManager_Payload"
#endif

[Setup]
AppId={{D4D46FC4-4F2E-4DC8-8C97-78C60392A822}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\RoomManager
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputBaseFilename=RoomManagerSetup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
SetupIconFile=
UninstallDisplayIcon={commonappdata}\Autodesk\Revit\Addins\2026\RoomManager.dll

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "runverify"; Description: "安装后运行依赖检查 (verify.bat)"; GroupDescription: "附加任务:"; Flags: unchecked

[Files]
Source: "{#PayloadDir}\*"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2026"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
Filename: "{cmd}"; Parameters: "/c ""{commonappdata}\Autodesk\Revit\Addins\2026\verify.bat"""; Description: "运行安装检查"; Flags: postinstall shellexec skipifsilent; Tasks: runverify

[UninstallDelete]
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\RoomManager.dll"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\RoomManager.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\ACadSharp.dll"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\EPPlus.dll"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\System.Drawing.Common.dll"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\verify.bat"

[Code]
function IsRevitInstalled: Boolean;
begin
  Result := DirExists('C:\Program Files\Autodesk\Revit 2026');
end;

function IsRevitApiPresent: Boolean;
begin
  Result := FileExists('C:\Program Files\Autodesk\Revit 2026\RevitAPI.dll')
    and FileExists('C:\Program Files\Autodesk\Revit 2026\RevitAPIUI.dll');
end;

function InitializeSetup: Boolean;
var
  answer: Integer;
begin
  Result := True;

  if not IsRevitInstalled then
  begin
    answer := MsgBox(
      '未检测到 Revit 2026。' + #13#10 +
      '建议先安装 Revit 2026 再安装插件。' + #13#10#13#10 +
      '是否仍要继续安装？',
      mbConfirmation, MB_YESNO);
    if answer = IDNO then
      Result := False;
    Exit;
  end;

  if not IsRevitApiPresent then
  begin
    answer := MsgBox(
      '检测到 Revit，但未找到 RevitAPI.dll / RevitAPIUI.dll。' + #13#10 +
      '这通常意味着 Revit 安装不完整或版本不匹配。' + #13#10#13#10 +
      '是否仍要继续安装？',
      mbConfirmation, MB_YESNO);
    if answer = IDNO then
      Result := False;
  end;
end;
