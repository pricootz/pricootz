#include "CertificateInfo.iss"

#define MyAppName "Pricop PowerPoint Tools"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Pricop"
#define MyAppId "{{A8A9B27F-58F2-4CC5-ACD8-CF4AE4551399}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Pricop\PowerPointTools
UsePreviousAppDir=no
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
VersionInfoVersion=2.0.0.0
VersionInfoCompany=Pricop
VersionInfoDescription=Pricop PowerPoint Tools VSTO Installer
VersionInfoProductName=Pricop PowerPoint Tools

[Files]
Source: "..\src\Pricop.PowerPointTools\bin\Release\Pricop.PowerPointTools.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\Pricop.PowerPointTools\bin\Release\Pricop.PowerPointTools.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\Pricop.PowerPointTools\bin\Release\Pricop.PowerPointTools.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\Pricop.PowerPointTools\bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\publisher.cer"; DestDir: "{app}"; Flags: ignoreversion

[InstallDelete]
Type: filesandordirs; Name: "{localappdata}\Pricop\PowerPointTools"

[Registry]
; New VSTO identity. We intentionally do NOT reuse the old COM ProgID.
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: string; ValueName: "Description"; ValueData: "PowerPoint productivity tools"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\Pricop.PowerPointTools.vsto|vstolocal"; Flags: uninsdeletekey

Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Pricop PowerPoint Tools"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: string; ValueName: "Description"; ValueData: "PowerPoint productivity tools"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools.VSTO"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\Pricop.PowerPointTools.vsto|vstolocal"; Flags: uninsdeletekey; Check: IsWin64

[Run]
; The build certificate is self-signed, so trust it explicitly for VSTO/ClickOnce.
Filename: "{sys}\certutil.exe"; Parameters: "-addstore -f Root ""{app}\publisher.cer"""; Flags: runhidden waituntilterminated
Filename: "{sys}\certutil.exe"; Parameters: "-addstore -f TrustedPublisher ""{app}\publisher.cer"""; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "{sys}\certutil.exe"; Parameters: "-delstore Root {#PublisherThumbprint}"; Flags: runhidden waituntilterminated
Filename: "{sys}\certutil.exe"; Parameters: "-delstore TrustedPublisher {#PublisherThumbprint}"; Flags: runhidden waituntilterminated

[Code]
procedure RemoveOldComAddin();
begin
  SetRegView(32);
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\Pricop.PowerPointTools');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\CLSID\{1D8EA74E-3C7F-4DA2-BE72-5C4FCF17A061}');
  if IsWin64 then
  begin
    SetRegView(64);
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Microsoft\Office\PowerPoint\Addins\Pricop.PowerPointTools');
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\Pricop.PowerPointTools');
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\CLSID\{1D8EA74E-3C7F-4DA2-BE72-5C4FCF17A061}');
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    Exec('taskkill.exe', '/IM POWERPNT.EXE /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    RemoveOldComAddin();
  end;
end;
