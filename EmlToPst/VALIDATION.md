# 검증 기록

작성일: 2026-11-05

## 실제로 수행한 검사

| 검사 | 결과 | 범위 |
| --- | --- | --- |
| XAML XML 및 코드 연결 | 통과 | App, MainWindow, BackupView, PstImportView 4개 파일과 대응 partial class |
| 화면 Binding 참조 | 통과 | 63개 Binding의 ViewModel/행 모델 속성 참조 및 이벤트 핸들러 |
| 프로젝트 XML/참조 | 통과 | 메인·회귀 검사 프로젝트의 패키지 버전 일치, 링크 소스 존재 |
| C# 어휘 구조 검사 | 통과 | 31개 C# 파일의 주석·문자열·보간 문자열을 고려한 괄호/블록 대응. 컴파일 검사는 아님 |
| 실제 SQLite 스키마·쿼리 | 통과 | 코드에서 추출한 스키마와 6개 변경 쿼리를 Python SQLite로 실행 |
| SQLite 상태 시나리오 | 통과 | pending/staged/complete, staged 실패 보존, 해시 없는 완료 거부, 재시도 시 초기화·횟수 증가, 메타데이터 갱신 |
| SQLite 식별/무결성 | 통과 | Int32를 넘는 UID, UIDVALIDITY 세대 분리, 파라미터에 SQL 형태의 폴더명 전달, integrity_check |
| OfficeIMO API 확인 | 확인 | 첨부한 실제 3.5.3 DLL의 메타데이터에서 MessageMetadata 접근, IsRead/IsDraft setter, 요약 IsRead getter 확인 |
| 로그 소유/통신 코드 검사 | 확인 | 작업별 로그 생성·Dispose 제거, 앱 시작/종료 소유, 읽기 전용 IMAP, 인증서 검증 우회 및 서버 변경 명령 없음 |

SQL 검사는 실제 C# 쿼리 문자열을 SQLite 엔진에 실행한 결과입니다. Microsoft.Data.Sqlite의 C# 바인딩, 네이티브 DLL 로딩, WPF 실행을 대신하는 검사는 아닙니다.

## 이 환경에서 수행하지 못한 검사

- `dotnet build EmlToPst.csproj -c Release -p:Platform=x64` 시도: `dotnet: command not found`.
- .NET SDK 및 NuGet 직접 다운로드도 현재 실행 환경의 네트워크 제한으로 사용할 수 없었습니다.
- 따라서 **C# 컴파일 성공, WPF 화면 실행, 13개 C# 회귀 시나리오의 실행 성공을 확인한 상태가 아닙니다.**
- 회사 IMAP 서버 로그인, 인증서, 계정 정책은 실제 회사 PC/계정으로 확인해야 합니다. 계정 비밀번호를 받거나 서버에 로그인하지 않았습니다.
- 클래식 Outlook의 PST 연결·표시도 실제 Windows/Outlook에서 확인해야 합니다.

## PC에서 실행할 검사

```powershell
.\Build.ps1
```

스크립트는 메인 WPF 프로젝트를 빌드한 다음 `Tests/BackupRegression`을 실행합니다. 회귀 검사는 회사 서버를 사용하지 않습니다. 실패한 단계가 있으면 스크립트가 오류로 종료합니다.

13개 시나리오는 최초/반복 백업, 동일 크기 손상 복구, 취소 후 재개, 부분 전송 후 재접속/비밀번호 마스킹, 재시도 상한, staged 임시/최종 파일 복구, UIDVALIDITY 세대 변경/재접속 중 변경, 서버 삭제, 경로 충돌/범위 검사, 미래 DB 버전 보호, 동시 로그, PST 수신 일자·읽음·첨부 바이트 왕복 검사입니다.

실제 로그인 후에는 작은 메일함부터 EML·PST 내용을 비교하고, 변환 없이 실행/종료했을 때에도 로그 첫 줄 `APP_START`, 마지막 줄 `APP_EXIT`가 있는지 확인하세요.
