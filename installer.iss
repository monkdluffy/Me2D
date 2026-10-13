[Setup]
AppName=Canido Doc Manager
AppVersion=1.0
AppPublisher=Canido
DefaultDirName={autopf}\Canido Doc Manager
DefaultGroupName=Canido Doc Manager
OutputBaseFilename=Install_Canido_Doc_Manager
Compression=lzma2
SolidCompression=yes
OutputDir=dist_installer
SetupIconFile=assets\logo.ico

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main Executable
Source: "dist\Canido Doc Manager.exe"; DestDir: "{app}"; Flags: ignoreversion

; Master Folders
Source: "Master SMAP\*"; DestDir: "{app}\Master SMAP"; Flags: ignoreversion recursesubdirs createallsubdirs

; Excel Data Batch
Source: "Template_Batch_Data.xlsx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Canido Doc Manager"; Filename: "{app}\Canido Doc Manager.exe"
Name: "{autodesktop}\Canido Doc Manager"; Filename: "{app}\Canido Doc Manager.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Canido Doc Manager.exe"; Description: "{cm:LaunchProgram,Canido Doc Manager}"; Flags: nowait postinstall skipifsilent
