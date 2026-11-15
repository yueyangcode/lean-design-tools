#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif

#define AppName "Lean Design Tools"
#define AppPublisher "Yueyang"
#define AddInDll "LeanDesignTools.dll"
#define AddInClsid "{5B4630A0-E087-4E24-93B5-8A79B0E1B842}"
#define PaneClsid "{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}"
#define ComCategory "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define AssemblyName "LeanDesignTools, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null"

[Setup]
AppId={{5B4630A0-E087-4E24-93B5-8A79B0E1B842}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\Lean Design Tools
DefaultGroupName=Lean Design Tools
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputDir=..\dist
OutputBaseFilename=LeanDesignTools-Setup-x64-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no
UninstallDisplayName=Lean Design Tools {#AppVersion}
VersionInfoVersion={#AppVersion}.0
VersionInfoCompany={#AppPublisher}
VersionInfoDescription=PowerPoint design tools installer
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}

[Files]
Source: "..\src\LeanDesignTools\bin\Release\net48\{#AddInDll}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\Classes\LeanDesignTools.AddIn"; ValueType: string; ValueData: "LeanDesignTools.AddIn"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\LeanDesignTools.AddIn\CLSID"; ValueType: string; ValueData: "{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}"; ValueType: string; ValueData: "LeanDesignTools.AddIn"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "LeanDesignTools.AddIn"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{code:GetCodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "LeanDesignTools.AddIn"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{code:GetCodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\ProgId"; ValueType: string; ValueData: "LeanDesignTools.AddIn"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{5B4630A0-E087-4E24-93B5-8A79B0E1B842}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"

Root: HKCU; Subkey: "Software\Classes\LeanDesignTools.Pane"; ValueType: string; ValueData: "LeanDesignTools.DesignPane"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\LeanDesignTools.Pane\CLSID"; ValueType: string; ValueData: "{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}"; ValueType: string; ValueData: "LeanDesignTools.DesignPane"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "LeanDesignTools.DesignPane"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{code:GetCodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "LeanDesignTools.DesignPane"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\InprocServer32\0.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{code:GetCodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\ProgId"; ValueType: string; ValueData: "LeanDesignTools.Pane"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\Control"

Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LeanDesignTools.AddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "设计工具（精简版）"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LeanDesignTools.AddIn"; ValueType: string; ValueName: "Description"; ValueData: "独立的 PowerPoint 设计工具任务窗格"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LeanDesignTools.AddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Code]
function GetCodeBase(Param: String): String;
begin
  Result := ExpandConstant('{app}\{#AddInDll}');
  StringChangeEx(Result, '\', '/', True);
end;

function KnownOfficePlatform(): String;
begin
  Result := '';
  RegQueryStringValue(HKEY_LOCAL_MACHINE_64,
    'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Result);
  if Result = '' then
    RegQueryStringValue(HKEY_LOCAL_MACHINE_32,
      'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Result);
end;

function InitializeSetup(): Boolean;
var
  Platform: String;
  Release: Cardinal;
begin
  Result := False;
  Platform := Lowercase(KnownOfficePlatform());
  if Platform = 'x86' then
  begin
    MsgBox('检测到 32 位 Office。此安装包仅支持 64 位 PowerPoint。', mbError, MB_OK);
    Exit;
  end;

  if (not RegQueryDWordValue(HKEY_LOCAL_MACHINE_64,
      'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release)) or
     (Release < 528040) then
  begin
    MsgBox('需要 Microsoft .NET Framework 4.8 或更高版本。', mbError, MB_OK);
    Exit;
  end;
  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if FindWindowByClassName('PPTFrameClass') <> 0 then
    Result := '请保存工作并关闭 PowerPoint，然后重新运行安装程序。';
end;

