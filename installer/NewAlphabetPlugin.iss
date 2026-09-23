; Builds NewAlphabetPlugin-Setup.exe, which installs the Word add-in for the current Windows user without asking for
; administrator rights. The build passes the version and the public key of the certificate that signed the add-in's
; manifests, for example:
;   iscc /DAppVersion=0.3.0 /DFileVersion=0.3.0.57 "/DPublicKey=<RSAKeyValue>...</RSAKeyValue>" NewAlphabetPlugin.iss

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef FileVersion
  #define FileVersion AppVersion
#endif
#ifndef SourceDir
  #define SourceDir "..\NewAlphabetPlugin\bin\Release"
#endif
#ifndef PublicKey
  #error Pass the public key the add-in's manifests are signed with: /DPublicKey=<RSAKeyValue>...</RSAKeyValue>
#endif

; Word finds the add-in through this key. It has the name Visual Studio gives the add-in, so an installed copy and a
; development build replace each other instead of both loading.
#define AddinKey "Software\Microsoft\Office\Word\Addins\NewAlphabetPlugin"

; An entry in the VSTO runtime's list of trusted add-ins, made the way Visual Studio makes them for its own builds.
; It trusts the manifests at this address that are signed with this key, so Word loads the add-in without asking.
; The entry's name never changes, so each update replaces it.
#define InclusionKey "Software\Microsoft\VSTO\Security\Inclusion\eb76c4ee-a1c9-4d9f-88c9-24ad235d5b23"

[Setup]
AppId={{6A9EBFE9-69DF-4845-BFE7-5AEDAF1DC681}
AppName=Yangi alifbo
AppVersion={#AppVersion}
AppPublisher=Diyorbek Satimboyev
AppPublisherURL=https://github.com/discoveruz/NewAlphabetPlugin
AppSupportURL=https://github.com/discoveruz/NewAlphabetPlugin/issues
AppUpdatesURL=https://github.com/discoveruz/NewAlphabetPlugin/releases
VersionInfoVersion={#FileVersion}
; For the current user only: %LOCALAPPDATA%\Programs\NewAlphabetPlugin, no administrator rights.
PrivilegesRequired=lowest
DefaultDirName={autopf}\NewAlphabetPlugin
DisableDirPage=yes
DisableProgramGroupPage=yes
UninstallDisplayName=Yangi alifbo (Word)
; An update replaces files that Word may have open, so setup offers to close Word first.
CloseApplications=yes
OutputDir=Output
OutputBaseFilename=NewAlphabetPlugin-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes

[Messages]
FinishedLabelNoIcons=Setup has finished installing [name]. The Yangi alifbo tab appears next to Home the next time Word opens.

; The add-in's only network use that the user has not just asked for, so it is shown here and can be turned off.
; See "Privacy" in README.md. An update keeps the earlier choice.
[Tasks]
Name: "updatecheck"; GroupDescription: "Updates:"; Description: "Look for new versions once a day when Word starts. This asks GitHub for the newest release; no documents or personal data are sent. Without it, the Ma'lumot button still looks when you press it."

[Files]
; The manifests are signed over the exact files, so they are installed exactly as built.
Source: "{#SourceDir}\NewAlphabetPlugin.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\NewAlphabetPlugin.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\NewAlphabetPlugin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; "|vstolocal" makes Word run the add-in from the installed folder rather than from ClickOnce's cache.
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:ManifestUrl}|vstolocal"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Yangi alifbo"
Root: HKCU; Subkey: "{#AddinKey}"; ValueType: string; ValueName: "Description"; ValueData: "Writes Uzbek text in the new Latin alphabet"
Root: HKCU; Subkey: "{#InclusionKey}"; ValueType: string; ValueName: "Url"; ValueData: "{code:ManifestUrl}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#InclusionKey}"; ValueType: string; ValueName: "PublicKey"; ValueData: "{#PublicKey}"
; Read by Updater.cs.
Root: HKCU; Subkey: "Software\NewAlphabetPlugin"; ValueType: dword; ValueName: "CheckForUpdates"; ValueData: 1; Tasks: updatecheck; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\NewAlphabetPlugin"; ValueType: dword; ValueName: "CheckForUpdates"; ValueData: 0; Tasks: not updatecheck; Flags: uninsdeletekey

[Code]
const
  VstoRuntimeUrl = 'https://go.microsoft.com/fwlink/?LinkId=158918';
  DotNetUrl = 'https://dotnet.microsoft.com/download/dotnet-framework/net48';

// The installed deployment manifest as a file:/// address, as Visual Studio writes it.
function ManifestUrl(Param: String): String;
begin
  Result := ExpandConstant('{app}\NewAlphabetPlugin.vsto');
  StringChangeEx(Result, '\', '/', True);
  Result := 'file:///' + Result;
end;

function HasVersionValue(RootKey: Integer; SubKey: String): Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(RootKey, SubKey, 'Version', Version) and (Version <> '');
end;

// The Visual Studio Tools for Office runtime, which runs the add-in inside Word. Office installs it under "v4", and
// Microsoft's own download under "v4R". The same checks as Microsoft's ClickOnce setup.
function VstoRuntimeInstalled: Boolean;
begin
  Result := HasVersionValue(HKLM32, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R')
    or HasVersionValue(HKLM32, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4');
  if not Result and IsWin64 then
    Result := HasVersionValue(HKLM64, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R')
      or HasVersionValue(HKLM64, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4');
end;

function InitializeSetup: Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;

  if not IsDotNetInstalled(net48, 0) then
  begin
    if SuppressibleMsgBox('Yangi alifbo needs .NET Framework 4.8, which is not installed.' + #13#10#13#10 +
      'Open its download page now? Run this setup again once it is installed.',
      mbError, MB_YESNO, IDNO) = IDYES then
      ShellExec('open', DotNetUrl, '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
    Result := False;
    Exit;
  end;

  if not VstoRuntimeInstalled then
  begin
    if SuppressibleMsgBox('Yangi alifbo needs the Microsoft Visual Studio Tools for Office runtime, which runs ' +
      'add-ins inside Word. It is not installed.' + #13#10#13#10 +
      'Download it from Microsoft now? Install it, then run this setup again.',
      mbError, MB_YESNO, IDNO) = IDYES then
      ShellExec('open', VstoRuntimeUrl, '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
    Result := False;
  end;
end;
