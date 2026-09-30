#define MyAppName "Pricop PowerPoint Tools"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "Pricop"
#define MyAppId "{{A8A9B27F-58F2-4CC5-ACD8-CF4AE4551399}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Pricop\PowerPointTools
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=..\dist
OutputBaseFilename=PricopPowerPointTools-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\app.ico
ArchitecturesAllowed=x86 x64compatible
CloseApplications=yes
RestartApplications=no
VersionInfoVersion=1.0.1.0
VersionInfoCompany=Pricop
VersionInfoDescription=Pricop PowerPoint Tools Installer
VersionInfoProductName=Pricop PowerPoint Tools

[Files]
Source: "..\src\Pricop.PowerPointTools\bin\Release\net48\Pricop.PowerPointTools.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Office add-in registration. Write both views on 64-bit Windows.
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "Description"; ValueData: "Productivity tools for PowerPoint"
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "Description"; ValueData: "Productivity tools for PowerPoint"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64

[Run]
; Use Microsoft's real .NET COM registration instead of hand-written CLSID keys.
Filename: "{win}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"; Parameters: """{app}\Pricop.PowerPointTools.dll"" /codebase"; Flags: runhidden waituntilterminated
Filename: "{win}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"; Parameters: """{app}\Pricop.PowerPointTools.dll"" /codebase"; Flags: runhidden waituntilterminated; Check: IsWin64

[UninstallRun]
Filename: "{win}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"; Parameters: """{app}\Pricop.PowerPointTools.dll"" /unregister"; Flags: runhidden waituntilterminated; Check: IsWin64
Filename: "{win}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"; Parameters: """{app}\Pricop.PowerPointTools.dll"" /unregister"; Flags: runhidden waituntilterminated

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
    Exec('taskkill.exe', '/IM POWERPNT.EXE /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
