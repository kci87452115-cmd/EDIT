; ============================================================
;  편집 프로그램 (RADIO / TV 런처) 설치 파일 - Inno Setup 6
;  GitHub Actions 가 /DAppVer /DTag /DSrcDir /DPkgDir 를 넘겨 컴파일합니다.
;
;  설치 결과 (C:\edit)
;    런처\EditLauncher.exe          런처 + 공용 실행 환경 (RADIO·TV 같이 씀)
;    업데이트\                      프로그램 꾸러미 (런처에서 필요한 것만 설치)
;    radio\앱 · 데이터 · 백업        TV\앱 · 데이터 · 백업   (데이터·백업은 제거해도 남음)
;    바탕화면 바로가기: [편집 프로그램] 하나
; ============================================================
#ifndef AppVer
  #define AppVer "0.0.0"
#endif
#ifndef SrcDir
  #define SrcDir "..\dist\EditLauncher"
#endif
#ifndef PkgDir
  #define PkgDir "..\Output\update"
#endif
#ifndef Tag
  #define Tag "dev"
#endif

[Setup]
; 예전 '운행표 데이터 편집기'(C:\RADIO) 설치와 같은 AppId -> 그 등록을 이어받아 깔끔하게 바뀐다. 절대 바꾸지 마세요.
AppId={{52764615-1BDF-4279-88BB-1732F176B599}
AppName=편집 프로그램 (RADIO / TV)
AppVersion={#AppVer}
AppVerName=편집 프로그램 (런처 {#AppVer})
DefaultDirName=C:\edit
; 예전 설치 위치(C:\RADIO)를 이어받지 않고 항상 C:\edit
UsePreviousAppDir=no
DisableDirPage=yes
DefaultGroupName=편집 프로그램
DisableProgramGroupPage=yes
UsePreviousGroup=no
PrivilegesRequired=lowest
OutputDir=..\Output
OutputBaseFilename=EDIT_setup_{#Tag}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
AppMutex=EditLauncher_Running,RADIOLauncher_Running,RunSheetEditor_Running
CloseApplications=yes
UninstallDisplayIcon={app}\런처\EditLauncher.exe

[Languages]
#if FileExists(AddBackslash(CompilerPath) + "Languages\Korean.isl")
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
#else
Name: "english"; MessagesFile: "compiler:Default.isl"
#endif

[Dirs]
Name: "{app}\업데이트"; Flags: uninsneveruninstall
Name: "{app}\radio\앱"
Name: "{app}\radio\데이터"; Flags: uninsneveruninstall
Name: "{app}\radio\백업"; Flags: uninsneveruninstall
Name: "{app}\TV\앱"
Name: "{app}\TV\데이터"; Flags: uninsneveruninstall
Name: "{app}\TV\백업"; Flags: uninsneveruninstall

[InstallDelete]
; 런처 폴더는 통째로 새로
Type: filesandordirs; Name: "{app}\런처"
; 예전 설치(C:\RADIO) 의 프로그램·바로가기 정리 - 예전 데이터는 그대로 (런처가 처음 켤 때 옮길지 묻는다)
Type: filesandordirs; Name: "C:\RADIO\프로그램"
Type: files; Name: "C:\RADIO\unins000.*"
Type: files; Name: "{userdesktop}\운행표 편집기.lnk"
Type: files; Name: "{userprograms}\운행표 편집기.lnk"
Type: filesandordirs; Name: "{userprograms}\운행표 편집기"

[Files]
Source: "{#SrcDir}\*"; DestDir: "{app}\런처"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#PkgDir}\catalog.json"; DestDir: "{app}\업데이트"; Flags: ignoreversion
Source: "{#PkgDir}\app_*.zip"; DestDir: "{app}\업데이트"; Flags: ignoreversion
Source: "기본데이터.db"; DestDir: "{app}\radio\데이터"; Flags: onlyifdoesntexist uninsneveruninstall skipifsourcedoesntexist

[Icons]
Name: "{userdesktop}\편집 프로그램"; Filename: "{app}\런처\EditLauncher.exe"; WorkingDir: "{app}\런처"
Name: "{group}\편집 프로그램"; Filename: "{app}\런처\EditLauncher.exe"; WorkingDir: "{app}\런처"
Name: "{group}\edit 폴더 열기"; Filename: "{app}"
Name: "{group}\편집 프로그램 제거"; Filename: "{uninstallexe}"

[UninstallDelete]
; 런처가 설치한 프로그램 파일 (데이터는 남김)
Type: filesandordirs; Name: "{app}\radio\앱"
Type: filesandordirs; Name: "{app}\TV\앱"

[Run]
Filename: "{app}\런처\EditLauncher.exe"; Description: "편집 프로그램 실행"; WorkingDir: "{app}\런처"; Flags: nowait postinstall
