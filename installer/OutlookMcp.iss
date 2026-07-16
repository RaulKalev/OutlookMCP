#ifndef MyAppVersion
  #define MyAppVersion "1.1.0"
#endif
#ifndef SourceRoot
  #define SourceRoot ".."
#endif
#ifndef OutputRoot
  #define OutputRoot "..\artifacts\installer"
#endif

[Setup]
AppId={{8AD98EF3-9ED2-4AE4-A20D-C6C0F03A6B6B}
AppName=EULE Outlook MCP
AppVersion={#MyAppVersion}
AppPublisher=EULE
DefaultDirName={localappdata}\Programs\EULE Outlook MCP
DefaultGroupName=EULE Outlook MCP
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputRoot}
OutputBaseFilename=EULE-Outlook-MCP-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
UninstallDisplayIcon={app}\OutlookMcp.Server.exe
WizardStyle=modern

[Files]
Source: "{#SourceRoot}\artifacts\publish\win-x64\OutlookMcp.Server.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\artifacts\publish\win-x64\config.sample.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\artifacts\publish\win-x64\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\artifacts\publish\win-x64\examples\*"; DestDir: "{app}\examples"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Outlook MCP Diagnostics"; Filename: "{app}\OutlookMcp.Server.exe"; Parameters: "--diagnose"; WorkingDir: "{app}"
Name: "{group}\Outlook MCP README"; Filename: "{app}\README.md"
Name: "{group}\Uninstall EULE Outlook MCP"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\OutlookMcp.Server.exe"; Parameters: "--diagnose"; Description: "Run Outlook MCP diagnostics"; Flags: postinstall nowait skipifsilent
