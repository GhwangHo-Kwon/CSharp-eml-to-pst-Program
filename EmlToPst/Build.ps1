param([switch]$Publish)
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
        throw '.NET 10 SDK를 설치한 뒤 다시 실행해 주세요.'
    }
    & dotnet build .\EmlToPst.csproj -c Release -p:Platform=x64
    if ($LASTEXITCODE -ne 0) { throw 'WPF 빌드 실패' }
    & dotnet run --project .\Tests\BackupRegression\BackupRegression.csproj -c Release
    if ($LASTEXITCODE -ne 0) { throw '회귀 검사 실패' }
    if ($Publish) {
        & dotnet publish .\EmlToPst.csproj -c Release -p:Platform=x64 -r win-x64 --self-contained false
        if ($LASTEXITCODE -ne 0) { throw '게시 폴더 생성 실패' }
    }
}
finally { Pop-Location }
