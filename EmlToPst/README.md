# EML → Outlook PST

Windows용 C# WPF MVVM 프로젝트입니다. 선택한 폴더의 EML을 새 Unicode PST 하나에 저장합니다. 하위 폴더 포함이 기본값이며, PST의 `EML 가져오기` 폴더 아래에 원본의 상대 폴더 구조를 재현합니다.

## 실행 환경

- 빌드에는 **.NET 10 SDK**가 필요합니다. Visual Studio를 사용할 경우 **Visual Studio 2026**에서 **.NET 데스크톱 개발** 워크로드를 설치합니다. [Microsoft의 .NET 설치 안내](https://learn.microsoft.com/dotnet/core/install/windows)
- EML 읽기와 PST 생성에는 MIT 라이선스의 `OfficeIMO.Email` 3.5.3을 사용합니다. PST 생성은 Outlook 설치 없이 가능합니다. [OfficeIMO.Email NuGet](https://www.nuget.org/packages/OfficeIMO.Email) · [OfficeIMO PST 문서](https://officeimo.com/docs/email/mailbox-stores/)
- 자동 연결을 사용할 때는 클래식 Outlook을 설치하고 기본 프로필을 설정합니다. Outlook 연결은 설치된 클래식 Outlook의 COM 개체 모델을 사용합니다. [Outlook AddStoreEx](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.addstoreex)

## 빌드 및 사용

1. ZIP을 풀고 `EmlToPst.sln`을 열어 **F5**를 누릅니다. Windows 터미널에서 실행할 수도 있습니다.

   ```powershell
   dotnet run --project .\EmlToPst.csproj
   ```

2. **폴더 선택**에서 EML 원본 폴더를 선택합니다. ZIP 백업을 사용하는 경우 먼저 압축을 풉니다.
3. **저장 위치**에서 새 `.pst` 파일명을 지정합니다.
4. **완료 후 Outlook에 PST 연결**을 선택하면 PST 저장이 끝난 뒤 클래식 Outlook에 별도 데이터 파일로 연결합니다. 기본값인 체크 해제 상태에서는 PST만 저장합니다.
5. **변환 시작**을 누릅니다. 저장·실패·미처리 건수와 PST 경로를 확인할 수 있습니다. 실패 목록에서 파일을 선택하고 `Ctrl+C`로 복사할 수 있습니다.

나중에 열려면 클래식 Outlook에서 **파일 → 열기 및 내보내기 → Outlook 데이터 파일 열기**를 선택합니다. [Microsoft의 PST 열기 안내](https://support.microsoft.com/ko-kr/outlook/open-and-close-outlook-data-files-pst)

## 변환 로그

변환을 시작할 때마다 `%LOCALAPPDATA%\EmlToPst\Logs\`에 별도의 UTF-8 `.log` 파일을 만듭니다. 앱에서 경로를 복사하거나 **로그 열기**로 메모장에서 확인할 수 있습니다. 로그 파일은 변환 중에도 열 수 있습니다.

| 로그 단계 | 기록 내용 |
| --- | --- |
| `START`, `SCAN` | 원본과 결과 경로, 옵션, 파일 검색 건수 |
| `EML_READ`, `PST_ADD_ITEM` | 파일별 상대 경로, 처리 순번, 누적 기록 및 읽기 실패 건수 |
| `PST_FINALIZE` | PST 저장 시작·완료, 저장 크기, 폴더와 메일 수, 변환 진단 코드 |
| `OUTLOOK_ATTACH` | 옵션 해제 등으로 건너뛴 이유 또는 Outlook 연결 결과 |
| `CANCEL`, `END` | 취소 시점, 최종 집계, 소요 시간, PST 경로 |

실패할 때는 관련 단계, 예외 메시지, HRESULT와 호출 스택을 기록합니다. 로그에는 원본 경로와 파일명이 포함됩니다. 변환 중 로그 기록이 불가능해지면 화면에 표시하고 PST 변환을 계속합니다.

## 동작과 복구

- 원본 EML은 읽기 전용으로 엽니다. 기존 PST를 덮어쓰지 않도록 하고, 정상 완료된 새 PST를 최종 경로에 확정합니다.
- 읽을 수 없는 EML은 실패로 기록하고 다음 파일로 넘어갑니다. PST 기록 또는 확정 단계에서 예외가 발생하면 변환을 중단하고 저장 여부를 표시합니다.
- 취소를 요청하면 현재 쓰기 작업을 마친 뒤 이미 기록한 메일을 PST로 저장합니다. 저장된 메일이 있더라도 취소한 작업에서는 Outlook 자동 연결을 건너뜁니다.
- 자동 연결에 실패해도 완성된 PST는 저장 위치에 남습니다.
- EML의 날짜, 인라인 이미지, 특수 인코딩, 서명·암호화 메일은 라이브러리가 변환하는 범위에 따릅니다. Outlook에서 표시되는 형태는 대표 샘플로 확인할 수 있습니다.

## 프로젝트 구조

| 위치 | 역할 |
| --- | --- |
| `Views/MainWindow.xaml` | 입력 선택, 진행률, 저장 결과, 로그 열기 |
| `ViewModels/MainViewModel.cs` | 화면 상태, 명령, 취소 |
| `Models/ImportModels.cs` | 옵션, 진행률, 결과 |
| `Services/PstImportService.cs` | EML 검색, 읽기, Unicode PST 쓰기, 저장 후 연결 분기 |
| `Services/OutlookPstAttacher.cs` | 이미 저장된 PST를 클래식 Outlook에 연결 |
| `Services/ImportLog.cs` | 변환 건별 로그 파일 생성 및 기록 |

EML 읽기와 PST 생성은 작업 스레드에서 수행합니다. 자동 연결을 선택했을 때만 WPF STA 스레드에서 Outlook COM을 호출합니다. [Outlook 개체 모델의 스레딩](https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook)

## 샘플 확인

`SampleEml`에는 가상 주소를 사용한 일반 텍스트 메일과 HTML 메일이 하나씩 있습니다. 아래 표의 시나리오마다 새 PST 파일명을 사용합니다.

| 확인 항목 | 예상 결과 |
| --- | --- |
| 클래식 Outlook을 닫고 자동 연결 해제 | Outlook을 시작하지 않고 PST 생성. 로그에는 `skipped: option disabled` 표시 |
| 하위 폴더 포함 | 메일 두 건 저장. 일반 메일은 `EML 가져오기`, HTML 메일은 `EML 가져오기\2026`에 표시 |
| 하위 폴더 옵션 해제 | 루트 EML 한 건만 저장 |
| 메일 원본 내용 | 한글 제목·본문, 보낸 사람·받는 사람, 날짜, 인라인 이미지와 `안내.txt` 첨부 확인 |
| 복사한 샘플 폴더에 0바이트 `empty.eml` 추가 | 메일 두 건 저장. 실패 목록 및 로그에 `empty.eml` 기록 |
| 기존 PST 경로를 다시 선택 | 기존 파일 유지. 로그에 경로 검증 실패 기록 |
| 자동 연결 선택 | 로그에 `PST_FINALIZE` 저장 완료가 `OUTLOOK_ATTACH` 시작보다 먼저 표시 |
| Outlook이 없는 PC에서 자동 연결 선택 | PST는 저장되고 화면과 로그에 연결 실패 표시 |

Windows 빌드와 실제 Outlook에서 PST 열기는 위 환경에서 확인해야 합니다.
