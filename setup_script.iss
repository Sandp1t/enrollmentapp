[Setup]
AppName=Enrollment App
AppVersion=1.0
DefaultDirName={autopf}\EnrollmentApp
DefaultGroupName=Enrollment App
OutputDir=output
OutputBaseFilename=EnrollmentAppInstaller
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\enrollment.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Enrollment App"; Filename: "{app}\enrollment.exe"; IconFilename: "{app}\app_icon.ico"
Name: "{userdesktop}\Enrollment App"; Filename: "{app}\enrollment.exe"; IconFilename: "{app}\app_icon.ico"

[Run]
Filename: "{app}\enrollment.exe"; Description: "Launch Enrollment App"; Flags: nowait postinstall
