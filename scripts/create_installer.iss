; Inno Setup 6 script. Build through scripts\build_windows.ps1, which passes
; /DBundleDir=<PyInstaller onedir output>. The file is saved as UTF-8 with BOM
; so Simplified Chinese text is read correctly by ISCC.

#define AppName "多表格统一检索系统"
#define AppVersion "2.0.0"
#define AppExe "MultiTableSearch.exe"
#define AppPublisher "Forken"

[Setup]
AppId={{7C1F6A52-3D8B-4E71-9A4C-2B5E0D9F1A63}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\MultiTableSearch
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=MultiTableSearch_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
MinVersion=10.0.17763
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}
; 卸载默认保留用户数据（数据库、托管副本、备份）。

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务:"; Flags: unchecked

[Files]
Source: "{#BundleDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "启动 {#AppName}"; Flags: nowait postinstall skipifsilent
