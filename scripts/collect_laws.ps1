# 근로기준법 외 노동 관련 법령과 시행령 본문을 받는다. 인증키는 환경변수 LAW_OC.
$ErrorActionPreference = 'Stop'
$OC = $env:LAW_OC
if (-not $OC) { throw 'LAW_OC 환경변수가 없습니다.' }
$root = Split-Path -Parent $PSScriptRoot
$dir = Join-Path $root 'data\laws'
New-Item -ItemType Directory -Force $dir | Out-Null

function Call-Api($endpoint, $params) {
  $params['OC'] = $OC; $params['type'] = 'JSON'
  $q = ($params.GetEnumerator() | ForEach-Object { $_.Key + '=' + [uri]::EscapeDataString([string]$_.Value) }) -join '&'
  for ($i = 1; $i -le 4; $i++) {
    try { $b = (New-Object Net.WebClient).DownloadData("http://www.law.go.kr/DRF/$endpoint.do?$q"); Start-Sleep -Milliseconds 250; return [Text.Encoding]::UTF8.GetString($b) }
    catch { Start-Sleep -Seconds (2 * $i) }
  }
  throw "API 호출 실패: $endpoint"
}

$base = '최저임금법', '근로자퇴직급여 보장법', '산업재해보상보험법', '산업안전보건법', '남녀고용평등과 일ㆍ가정 양립 지원에 관한 법률',
        '기간제 및 단시간근로자 보호 등에 관한 법률', '파견근로자 보호 등에 관한 법률', '임금채권보장법', '노동조합 및 노동관계조정법',
        '근로자참여 및 협력증진에 관한 법률', '고용보험법'
$names = foreach ($b in $base) { $b; "$b 시행령" }
$index = @()
foreach ($name in $names) {
  $r = (Call-Api 'lawSearch' @{ target = 'law'; query = $name; display = '30' }) | ConvertFrom-Json
  $hit = @($r.LawSearch.law) | Where-Object { $_ -and $_.법령명한글 -eq $name } | Select-Object -First 1
  if (-not $hit) { Write-Host "없음: $name"; continue }
  $body = Call-Api 'lawService' @{ target = 'law'; MST = $hit.법령일련번호 }
  $file = "$($hit.법령ID).json"
  [IO.File]::WriteAllText((Join-Path $dir $file), $body, (New-Object Text.UTF8Encoding($false)))
  $index += [pscustomobject]@{ 법령명 = $name; 법령ID = $hit.법령ID; MST = $hit.법령일련번호; 시행일자 = $hit.시행일자; 파일 = $file }
  Write-Host "저장: $name (시행 $($hit.시행일자))"
}
[IO.File]::WriteAllText((Join-Path $dir 'index.json'), ($index | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($false)))
Write-Host "완료: $($index.Count)개 법령"
