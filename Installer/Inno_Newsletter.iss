; Installation of JAZZ live AARAU Newsletter

[Setup]
AppPublisher=JAZZ live AARAU
AppPublisherURL=https://jazzliveaarau.ch/
AppName=JAZZ live AARAU Newsletter
AppVerName=JAZZ live AARAU Newsletter version 3.20
DefaultDirName={sd}\Apps\JazzLiveAarau\Newsletter
DefaultGroupName=JAZZ live AARAU Newsletter
UninstallDisplayIcon={app}\SendNewsLetter.exe
Compression=lzma
SolidCompression=yes
OutputDir= NeueVersion
OutputBaseFilename= SetupJazzLiveAarauNewsletter-version-3-21

[Dirs]
Name: "{app}\Anhang"; Permissions: users-modify
Name: "{app}\Excel"; Permissions: users-modify
Name: "{app}\Plakat"; Permissions: users-modify
Name: "{app}\Help"; Permissions: users-modify
Name: "{app}\NeueVersion"; Permissions: users-modify
Name: "{app}\Vorlagen"; Permissions: users-modify
Name: "{app}\LatestVersionInfo"; Permissions: users-modify
Name: "{app}\XML"; Permissions: users-modify
Name: "{app}\EML"; Permissions: users-modify

[Files]
Source: "SendNewsLetter.exe"; DestDir: "{app}"
Source: "ExcelUtil.dll"; DestDir: "{app}"
Source: "Eml.dll"; DestDir: "{app}"
Source: "Ftp.dll"; DestDir: "{app}"
Source: "JazzFtp.dll"; DestDir: "{app}"
Source: "JazzApp.dll"; DestDir: "{app}"
Source: "JazzVersion.dll"; DestDir: "{app}"
Source: "EncodingTools.dll"; DestDir: "{app}"
Source: "Help\JAZZ_live_AARAU_Newsletter.rtf"; DestDir: "{app}\Help"; Flags: isreadme; Permissions: users-modify
Source: "Plakat\Logo.jpg";  DestDir: "{app}\Plakat"; Permissions: users-modify

[Icons]
Name: "{group}\JAZZ live AARAU Newsletter"; Filename: "{app}\SendNewsLetter.exe"

[InstallDelete]
Type: files; Name: "{app}\SendNewsletter.config"
Type: files; Name: "{app}\Eml.dll"
Type: files; Name: "{app}\Ftp.dll"
Type: files; Name: "{app}\JazzFtp.dll"
Type: files; Name: "{app}\ExcelUtil.dll"
Type: files; Name: "{app}\JazzApp.dll"
Type: files; Name: "{app}\JazzVersion.dll"
Type: files; Name: "{app}\EncodingTools.dll"

[UninstallDelete]
Type: files; Name: "{app}\SendNewsletter.config"
