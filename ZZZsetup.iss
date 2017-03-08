; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Blackd Safe Cheats"
#define MyAppVersion "2.4.0"
#define MyAppPublisher "blackdtools.com"
#define MyAppURL "http://blackdtools.com"
#define MyAppExeName "Tibia.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
PrivilegesRequired=admin
;Encryption=yes
;Password=vip123
AppId={{F4CFBC5D-12D5-423E-A4A3-BCB2F1631FD8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
LicenseFile=readme.rtf
OutputDir=_installer
OutputBaseFilename=bsc_installer_{#MyAppVersion}
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "catalan"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "czech"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "danish"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "finnish"; MessagesFile: "compiler:Languages\Finnish.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"
Name: "hungarian"; MessagesFile: "compiler:Languages\Hungarian.isl"
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "norwegian"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "serbiancyrillic"; MessagesFile: "compiler:Languages\SerbianCyrillic.isl"
Name: "serbianlatin"; MessagesFile: "compiler:Languages\SerbianLatin.isl"
Name: "slovenian"; MessagesFile: "compiler:Languages\Slovenian.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "Tibia.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-772.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-857.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-860.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-861.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-862.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-870.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-871.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-872.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-873.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-874.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-900.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-910.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-920.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-931.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-940.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-941.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-942.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-943.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-944.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-945.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-946.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-950.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-951.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-952.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-953.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-954.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-960.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-961.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-962.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-963.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-970.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-971.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-980.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-981.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-982.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-983.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-984.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-985.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-986.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-990.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-991.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-992.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-993.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1000.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1001.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1002.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1010.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1011.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1012.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1020.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1021.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1021preview.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1022.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1030.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1031.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1032.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1033.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1034.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1035.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1036.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1037.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1038.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1039.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1040.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1041.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1050.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1050preview.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1051.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1051preview.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1052.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1052preview.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1053.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1053preview.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1054.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1055.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1056.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1057.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1058.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1059.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1060.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1061.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1062.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1063.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1064.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1070.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1071.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1072.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1073.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1074.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1075.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1076.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1077.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1078.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1079.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1080.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1081.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1082.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1090.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1091.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1092.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1093.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1094.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1095.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1096.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1097.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1098.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1099-old0.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1099-old1.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1099-old2.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1099-old3.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1099-old4.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "conf-1099-old5.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.int"; DestDir: "{app}"; Flags: ignoreversion
Source: "default.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "lang_english.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "lang_polish.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "lang_portugues.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "lang_spanish.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "lang_swedish.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "lang_german.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "MSWINSCK.oca"; DestDir: "{app}"; Flags: 
Source: "MSWINSCK.OCX"; DestDir: "{app}"; Flags: 
Source: "readme.rtf"; DestDir: "{app}"; Flags: ignoreversion
Source: "Tibia.exe"; DestDir: "{app}"; Flags: ignoreversion
; [[[ begin VB6 system files
Source: "vbfiles\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "vbfiles\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vbfiles\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
; end VB6 system files ]]]
; [[[ begin custom additional VB6 files
Source: "mswinsck.ocx"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "mswsock.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
; end custom additional VB6 files ]]]
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

