[Setup]
AppName=Enrollment App
AppVersion=1.0
DefaultDirName={pf}\EnrollmentApp
DefaultGroupName=EnrollmentApp
OutputBaseFilename=EnrollmentAppInstaller
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\EnrollmentApp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Enrollment App"; Filename: "{app}\EnrollmentApp.exe"; IconFilename: "app_icon.ico"
Name: "{group}\Uninstall Enrollment App"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\EnrollmentApp.exe"; Description: "Launch Enrollment App"; Flags: nowait postinstall skipifsilent
