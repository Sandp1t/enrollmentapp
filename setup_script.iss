[Setup]
AppName=Enrollment App
AppVersion=1.0
DefaultDirName={autopf}\EnrollmentApp
DefaultGroupName=Enrollment App
OutputBaseFilename=EnrollmentAppInstaller
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest

[Files]
Source: "dist\enrollment.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"

[Icons]
Name: "{group}\Enrollment App"; Filename: "{app}\enrollment.exe"; IconFilename: "{app}\app_icon.ico"
Name: "{userdesktop}\Enrollment App"; Filename: "{app}\enrollment.exe"; IconFilename: "{app}\app_icon.ico"

[Run]
Filename: "{app}\enrollment.exe"; Description: "Launch Enrollment App"; Flags: nowait postinstall
