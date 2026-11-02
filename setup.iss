; ============================================================
;  운행표 데이터 편집기 설치 파일 (Inno Setup 6)
;  직접 컴파일하지 말고 release.py 가 버전을 넘겨 컴파일합니다.
;
;  설치 결과
;    C:\RADIO\프로그램\RunSheetEditor.exe   (업데이트 때 이 폴더만 교체)
;    C:\RADIO\데이터\                       (기록 DB·설정 - 업데이트/삭제해도 남음)
;    C:\RADIO\백업\                         (자동 백업 - 남음)
;    C:\RADIO\업데이트\                     (기본 업데이트 폴더 - 남음)
;    바탕화면·시작 메뉴 바로가기 "운행표 편집기"
; ============================================================
#ifndef AppVer
  #define AppVer "0.0.0"
#endif
#ifndef SrcDir
  #define SrcDir "dist\RunSheetEditor"
#endif
#define AppExe "RunSheetEditor.exe"

[Setup]
; AppId 는 절대 바꾸지 마세요 (바꾸면 업데이트가 아니라 다른 프로그램으로 설치됩니다)
AppId={{52764615-1BDF-4279-88BB-1732F176B599}
AppName=운행표 데이터 편집기
AppVersion={#AppVer}
AppVerName=운행표 데이터 편집기 {#AppVer}
DefaultDirName=C:\RADIO
DisableDirPage=yes
DefaultGroupName=운행표 편집기
DisableProgramGroupPage=yes
; 관리자 권한 없이 설치 -> 자동 업데이트 때 UAC 창이 뜨지 않고, 데이터 폴더 쓰기 권한 문제도 없다
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=RunSheetEditor_setup_{#AppVer}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; 프로그램이 켜져 있으면 닫으라고 안내 (프로그램 안의 APP_MUTEX 와 같은 이름)
AppMutex=RunSheetEditor_Running
CloseApplications=yes
UninstallDisplayIcon={app}\프로그램\{#AppExe}
#if FileExists("icon.ico")
SetupIconFile=icon.ico
#endif

[Languages]
#if FileExists(AddBackslash(CompilerPath) + "Languages\Korean.isl")
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
#else
Name: "english"; MessagesFile: "compiler:Default.isl"
#endif

[Dirs]
Name: "{app}\데이터"; Flags: uninsneveruninstall
Name: "{app}\백업"; Flags: uninsneveruninstall
Name: "{app}\업데이트"; Flags: uninsneveruninstall

[InstallDelete]
; 업데이트할 때 예전 프로그램 파일을 깨끗이 지우고 새로 넣는다 (데이터·백업 폴더는 건드리지 않음)
Type: filesandordirs; Name: "{app}\프로그램"

[Files]
Source: "{#SrcDir}\*"; DestDir: "{app}\프로그램"; Flags: ignoreversion recursesubdirs createallsubdirs
; 기본 데이터: 처음 설치할 때만 들어가고, 이미 있으면 절대 덮어쓰지 않는다
Source: "기본데이터.db"; DestDir: "{app}\데이터"; Flags: onlyifdoesntexist uninsneveruninstall skipifsourcedoesntexist

[Icons]
Name: "{userdesktop}\운행표 편집기"; Filename: "{app}\프로그램\{#AppExe}"; WorkingDir: "{app}\프로그램"
Name: "{group}\운행표 편집기"; Filename: "{app}\프로그램\{#AppExe}"; WorkingDir: "{app}\프로그램"
Name: "{group}\RADIO 폴더 열기"; Filename: "{app}"
Name: "{group}\운행표 편집기 제거"; Filename: "{uninstallexe}"

[Run]
; 설치가 끝나면 실행 (자동 업데이트(/SILENT) 때도 다시 켜진다)
Filename: "{app}\프로그램\{#AppExe}"; Description: "운행표 편집기 실행"; WorkingDir: "{app}\프로그램"; Flags: nowait postinstall
