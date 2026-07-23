#define MyAppName "Sarwar Smart PCMS"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Sarwar Smart PCMS"
#define MyAppExeName "SarwarSmartPCMS-AddIn64-packed.xll"

[Setup]
AppId={{59C12D1B-82E0-42C5-A913-7C6C3A7B9B01}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Sarwar Smart PCMS
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=Sarwar-Smart-PCMS-Setup-x86
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x86 x64
UninstallDisplayIcon={app}\SarwarSmartPCMS-AddIn-packed.xll

[Files]
Source: "..\src\bin\Release\net48\publish\SarwarSmartPCMS-AddIn-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\bin\Release\net48\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{userappdata}\Microsoft\Excel\XLSTART"

[InstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\SarwarSmartPCMS.xll"

[Run]
Filename: "{cmd}"; Parameters: "/C copy /Y ""{app}\SarwarSmartPCMS-AddIn-packed.xll"" ""{userappdata}\Microsoft\Excel\XLSTART\SarwarSmartPCMS.xll"""; Flags: runhidden waituntilterminated

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\SarwarSmartPCMS.xll"
