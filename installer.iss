#ifndef AppVersion
  #define AppVersion "1.7.0"
#endif

[Setup]
AppId={{83C8839C-702F-4B57-9238-3404BFB34E95}}
AppName=Rainy Teacher
AppVersion={#AppVersion}
AppPublisher=RAINY
AppPublisherURL=https://github.com/ssboroo/Rainy-Teacher
DefaultDirName={localappdata}\Programs\Rainy Teacher
DefaultGroupName=Rainy Teacher
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\RainyTeacher.exe
SetupIconFile=assets\rainy_teacher.ico
OutputDir=dist-installer
OutputBaseFilename=RainyTeacher-Setup-{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
Uninstallable=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional options:"; Flags: unchecked

[Files]
Source: "dist\RainyTeacher.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Rainy Teacher"; Filename: "{app}\RainyTeacher.exe"; IconFilename: "{app}\RainyTeacher.exe"
Name: "{autodesktop}\Rainy Teacher"; Filename: "{app}\RainyTeacher.exe"; IconFilename: "{app}\RainyTeacher.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\RainyTeacher.exe"; Description: "Launch Rainy Teacher"; Flags: postinstall nowait skipifsilent
