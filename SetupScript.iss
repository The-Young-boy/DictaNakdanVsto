[Setup]
AppName=נקדן דיקטה לוורד
AppVersion=1.0.0
AppPublisher=The Young Boy
DefaultDirName={autopf}\DictaNakdan
DefaultGroupName=Dicta Nakdan
OutputDir=userdocs:InnoSetupOutput
OutputBaseFilename=DictaNakdan_Installer
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\icon.ico
Compression=lzma
SolidCompression=yes
; הסרת התוסף מלוח הבקרה (Add/Remove Programs) תשתמש בסמל הזה

[Languages]
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"

[Files]
; === שנה את הנתיב הזה לנתיב של תיקיית ה-Publish שיצרת ב-Visual Studio ===
Source: "D:\Users\sdwac\source\repos\DictaNakdanVsto\DictaNakdanVsto\publish*"; DestDir: "{app}\InstallFiles"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
; מריץ את ההתקנה של VSTO שקופה ברקע כשההתקנה מסתיימת
Filename: "{commoncf64}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"; Parameters: "/install ""{app}\InstallFiles\DictaNakdanVsto.vsto"" /silent"; Flags: runhidden

[UninstallRun]
; מסיר את התוסף מהוורד בעת הסרת התוכנה
Filename: "{commoncf64}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"; Parameters: "/uninstall ""{app}\InstallFiles\DictaNakdanVsto.vsto"" /silent"; Flags: runhidden