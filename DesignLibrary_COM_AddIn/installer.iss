; ============================================================
; Design Library - PowerPoint COM Add-in Installer
; Built with Inno Setup (https://jrsoftware.org/isinfo.php)
; ============================================================

#define MyAppName "Design Library for PowerPoint"
#define MyAppVersion "1.0"
#define MyAppPublisher "Aveti Learning"

[Setup]
AppId={{CF8DBA7F-EDDE-4A8E-AF10-C3D7BB89EE69}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={commonpf}\DesignLibrary
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=DesignLibrary_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=compiler:SetupClassicIcon.ico
UninstallDisplayIcon={app}\uninstall.exe
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=POWERPNT.EXE

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Main Add-in DLL and Ribbon XML
Source: "DesignLibraryAddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "Ribbon.xml"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; --- PowerPoint Add-in Registration ---
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\DesignLibraryAddIn.AddIn"; ValueName: "Description"; ValueType: string; ValueData: "PowerPoint Design Library Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\DesignLibraryAddIn.AddIn"; ValueName: "FriendlyName"; ValueType: string; ValueData: "Design Library"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\DesignLibraryAddIn.AddIn"; ValueName: "LoadBehavior"; ValueType: dword; ValueData: "3"; Flags: uninsdeletekey

[Run]
; Register the COM DLL using regasm (64-bit .NET 4.x)
Filename: "{dotnet4064}\regasm.exe"; Parameters: """{app}\DesignLibraryAddIn.dll"" /codebase"; StatusMsg: "Registering COM Add-in..."; Flags: runhidden waituntilterminated
; Also register with 32-bit regasm for 32-bit Office
Filename: "{dotnet4032}\regasm.exe"; Parameters: """{app}\DesignLibraryAddIn.dll"" /codebase"; StatusMsg: "Registering COM Add-in (32-bit)..."; Flags: runhidden waituntilterminated skipifdoesntexist

[UninstallRun]
; Unregister the COM DLL on uninstall
Filename: "{dotnet4064}\regasm.exe"; Parameters: """{app}\DesignLibraryAddIn.dll"" /unregister"; Flags: runhidden waituntilterminated skipifdoesntexist
Filename: "{dotnet4032}\regasm.exe"; Parameters: """{app}\DesignLibraryAddIn.dll"" /unregister"; Flags: runhidden waituntilterminated skipifdoesntexist

[Code]
// Check if .NET Framework 4.x is installed
function IsDotNetInstalled(): Boolean;
var
  Version: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Version);
  if Result then
    Result := (Version >= 378389);
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not IsDotNetInstalled() then
  begin
    MsgBox('This add-in requires .NET Framework 4.5 or later.' + #13#10 +
           'Please install it from microsoft.com and try again.', mbError, MB_OK);
    Result := False;
  end;
end;

[Messages]
FinishedLabel=Setup has finished installing [name] on your computer.%n%nPlease restart PowerPoint to see the "Design Library" tab in the Ribbon.
