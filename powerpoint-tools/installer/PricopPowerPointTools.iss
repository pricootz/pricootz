#define MyAppName "Pricop PowerPoint Tools"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Pricop"
#define MyAppId "{{A8A9B27F-58F2-4CC5-ACD8-CF4AE4551399}"
#define AddinClsid "{1D8EA74E-3C7F-4DA2-BE72-5C4FCF17A061}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Pricop\PowerPointTools
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
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
VersionInfoVersion=1.0.0.0
VersionInfoCompany=Pricop
VersionInfoDescription=Pricop PowerPoint Tools Installer
VersionInfoProductName=Pricop PowerPoint Tools

[Files]
Source: "..\src\Pricop.PowerPointTools\bin\Release\net48\Pricop.PowerPointTools.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}"; ValueType: string; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "ThreadingModel"; ValueType: string; ValueData: "Both"
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "Class"; ValueType: string; ValueData: "Pricop.PowerPointTools.Connect"
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "Assembly"; ValueType: string; ValueData: "Pricop.PowerPointTools, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "RuntimeVersion"; ValueType: string; ValueData: "v4.0.30319"
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "CodeBase"; ValueType: string; ValueData: "file:///{app}\Pricop.PowerPointTools.dll"
Root: HKCU64; Subkey: "Software\Classes\Pricop.PowerPointTools"; ValueType: string; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Classes\Pricop.PowerPointTools\CLSID"; ValueType: string; ValueData: "{#AddinClsid}"

Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}"; ValueType: string; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "ThreadingModel"; ValueType: string; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "Class"; ValueType: string; ValueData: "Pricop.PowerPointTools.Connect"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "Assembly"; ValueType: string; ValueData: "Pricop.PowerPointTools, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "RuntimeVersion"; ValueType: string; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddinClsid}\InprocServer32"; ValueName: "CodeBase"; ValueType: string; ValueData: "file:///{app}\Pricop.PowerPointTools.dll"
Root: HKCU32; Subkey: "Software\Classes\Pricop.PowerPointTools"; ValueType: string; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\Pricop.PowerPointTools\CLSID"; ValueType: string; ValueData: "{#AddinClsid}"

Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "Description"; ValueData: "Productivity tools for PowerPoint"
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: string; ValueName: "Description"; ValueData: "Productivity tools for PowerPoint"
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
    Exec('taskkill.exe', '/IM POWERPNT.EXE /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
