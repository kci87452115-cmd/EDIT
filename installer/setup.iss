; ============================================================
;  RADIO 프로그램 (런처) 설치 파일 - Inno Setup 6
;  GitHub Actions 가 /DAppVer /DSrcDir /DPkgDir 를 넘겨 컴파일합니다.
;
;  설치 결과
;    C:\RADIO\런처\RADIOLauncher.exe   런처 + 공용 실행 환경
;    C:\RADIO\업데이트\                 프로그램 꾸러미 (런처에서 필요한 것만 설치)
;    C:\RADIO\앱 / 데이터 / 백업        (데이터·백업은 제거해도 남음)
;    바탕화면 바로가기: [RADIO 프로그램] 하나
; ============================================================
#ifndef AppVer
  #define AppVer "0.0.0"
#endif
#ifndef SrcDir
  #define SrcDir "..\dist\RADIOLauncher"
#endif
#ifndef PkgDir
  #define PkgDir "..\Output\update"
#endif
#ifndef Tag
  #define Tag "dev"
#endif

[Setup]
; 예전 '운행표 데이터 편집기' 설치와 같은 AppId -> 그 위에 덮어 설치되며 깔끔하게 바뀐다. 절대 바꾸지 마세요.
AppId={{52764615-1BDF-4279-88BB-1732F176B599}
AppName=RADIO 프로그램
AppVersion={#AppVer}
AppVerName=RADIO 프로그램 (런처 {#AppVer})
DefaultDirName=C:\RADIO
DisableDirPage=yes
DefaultGroupName=RADIO 프로그램
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\Output
OutputBaseFilename=RADIO_setup_{#Tag}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
AppMutex=RADIOLauncher_Running,RunSheetEditor_Running
CloseApplications=yes
UninstallDisplayIcon={app}\런처\RADIOLauncher.exe

[Languages]
#if FileExists(AddBackslash(CompilerPath) + "Languages\Korean.isl")
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
#else
Name: "english"; MessagesFile: "compiler:Default.isl"
#endif

[Dirs]
Name: "{app}\앱"
Name: "{app}\데이터"; Flags: uninsneveruninstall
Name: "{app}\백업"; Flags: uninsneveruninstall
Name: "{app}\업데이트"; Flags: uninsneveruninstall

[InstallDelete]
; 런처 폴더는 통째로 새로
Type: filesandordirs; Name: "{app}\런처"
; 예전 단독 설치(운행표 편집기) 정리 - 데이터는 그대로
Type: filesandordirs; Name: "{app}\프로그램"
Type: files; Name: "{userdesktop}\운행표 편집기.lnk"
Type: files; Name: "{userprograms}\운행표 편집기.lnk"
Type: filesandordirs; Name: "{userprograms}\운행표 편집기"

[Files]
Source: "{#SrcDir}\*"; DestDir: "{app}\런처"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#PkgDir}\catalog.json"; DestDir: "{app}\업데이트"; Flags: ignoreversion
Source: "{#PkgDir}\app_*.zip"; DestDir: "{app}\업데이트"; Flags: ignoreversion
Source: "기본데이터.db"; DestDir: "{app}\데이터"; Flags: onlyifdoesntexist uninsneveruninstall skipifsourcedoesntexist

[Icons]
Name: "{userdesktop}\RADIO 프로그램"; Filename: "{app}\런처\RADIOLauncher.exe"; WorkingDir: "{app}\런처"
Name: "{group}\RADIO 프로그램"; Filename: "{app}\런처\RADIOLauncher.exe"; WorkingDir: "{app}\런처"
Name: "{group}\RADIO 폴더 열기"; Filename: "{app}"
Name: "{group}\RADIO 프로그램 제거"; Filename: "{uninstallexe}"

[UninstallDelete]
; 런처가 설치한 프로그램 파일 (데이터는 남김)
Type: filesandordirs; Name: "{app}\앱"

[Run]
Filename: "{app}\런처\RADIOLauncher.exe"; Description: "RADIO 프로그램 실행"; WorkingDir: "{app}\런처"; Flags: nowait postinstall
