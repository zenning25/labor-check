# 국가법령정보 Open API 수집 스크립트. 인증키는 환경변수 LAW_OC에서 읽는다.
# 판례는 참조법령(JO)별 대법원 판례를 전부, 해석은 쟁점 키워드별로 끝까지 받는다. 이미 받은 본문은 건너뛴다.
$ErrorActionPreference = 'Stop'
$OC = $env:LAW_OC
if (-not $OC) { throw 'LAW_OC 환경변수가 없습니다.' }

$root = Split-Path -Parent $PSScriptRoot
$data = Join-Path $root 'data'
foreach ($d in 'law', 'prec', 'expc', 'byl') { New-Item -ItemType Directory -Force (Join-Path $data $d) | Out-Null }

function Save-Text($path, $text) {
  [IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding($false)))
}

function Call-Api($endpoint, $params) {
  $params['OC'] = $OC
  $params['type'] = 'JSON'
  $q = ($params.GetEnumerator() | ForEach-Object { $_.Key + '=' + [uri]::EscapeDataString([string]$_.Value) }) -join '&'
  for ($i = 1; $i -le 4; $i++) {
    try {
      $b = (New-Object Net.WebClient).DownloadData("http://www.law.go.kr/DRF/$endpoint.do?$q")
      Start-Sleep -Milliseconds 250
      return [Text.Encoding]::UTF8.GetString($b)
    } catch { Start-Sleep -Seconds (2 * $i) }
  }
  throw "API 호출 실패: $endpoint"
}

# 목록을 페이지 끝까지 받는다. $pick 은 응답 JSON에서 항목 배열과 총 건수를 꺼내는 스크립트블록.
function Get-All($params, $pick, $cap) {
  $all = New-Object Collections.Generic.List[object]
  $page = 1
  while ($true) {
    $p = $params.Clone(); $p['display'] = '100'; $p['page'] = [string]$page
    $r = (Call-Api 'lawSearch' $p) | ConvertFrom-Json
    $res = & $pick $r
    $items = @($res.items | Where-Object { $_ })
    foreach ($it in $items) { $all.Add($it) }
    if ($items.Count -lt 100 -or $all.Count -ge [int]$res.total -or $all.Count -ge $cap) { break }
    $page++
  }
  return $all
}

# 본문 파일을 4개씩 동시에 받는다. 이미 있는 파일은 호출하는 쪽에서 걸러 보낸다.
function Download-Many($jobs, $label) {
  $jobs = @($jobs)
  if (-not $jobs.Count) { Write-Host "${label}: 받을 것 없음"; return }
  $pool = [runspacefactory]::CreateRunspacePool(1, 4); $pool.Open()
  $sb = {
    param($url, $file)
    for ($i = 1; $i -le 4; $i++) {
      try { $b = (New-Object Net.WebClient).DownloadData($url); [IO.File]::WriteAllBytes($file, $b); Start-Sleep -Milliseconds 120; return 'ok' }
      catch { Start-Sleep -Seconds (2 * $i) }
    }
    return 'fail'
  }
  $done = 0; $fail = 0
  for ($i = 0; $i -lt $jobs.Count; $i += 100) {
    $chunk = @($jobs[$i..([Math]::Min($i + 99, $jobs.Count - 1))])
    $hs = foreach ($j in $chunk) {
      $ps = [powershell]::Create(); $ps.RunspacePool = $pool
      [void]$ps.AddScript($sb).AddArgument($j.url).AddArgument($j.file)
      [pscustomobject]@{ ps = $ps; ar = $ps.BeginInvoke() }
    }
    foreach ($h in $hs) { $r = $h.ps.EndInvoke($h.ar); if ("$r" -ne 'ok') { $fail++ }; $h.ps.Dispose() }
    $done += $chunk.Count; Write-Host "$label $done/$($jobs.Count) (실패 $fail)"
  }
  $pool.Close()
}

# 1. 법령 본문
$lawNames = '근로기준법', '근로기준법 시행령', '근로기준법 시행규칙'
$index = @()
foreach ($name in $lawNames) {
  $r = (Call-Api 'lawSearch' @{ target = 'law'; query = $name; display = '20' }) | ConvertFrom-Json
  $hit = @($r.LawSearch.law) | Where-Object { $_.법령명한글 -eq $name } | Select-Object -First 1
  if (-not $hit) { Write-Host "법령 없음: $name"; continue }
  $body = Call-Api 'lawService' @{ target = 'law'; MST = $hit.법령일련번호 }
  Save-Text (Join-Path $data "law\$($name -replace ' ','_').json") $body
  $index += [pscustomobject]@{ 법령명 = $name; MST = $hit.법령일련번호; 시행일자 = $hit.시행일자 }
  Write-Host "법령 저장: $name (시행 $($hit.시행일자))"
}
Save-Text (Join-Path $data 'law\index.json') ($index | ConvertTo-Json -Depth 4)

# 2. 별표 목록
Save-Text (Join-Path $data 'byl\list_시행령.json') (Call-Api 'lawSearch' @{ target = 'licbyl'; query = '근로기준법 시행령'; display = '50' })

# 3. 판례: 참조법령별 대법원 판례 전부
$joList = '근로기준법', '최저임금법', '근로자퇴직급여 보장법', '산업재해보상보험법', '산업안전보건법', '임금채권보장법',
          '기간제 및 단시간근로자 보호 등에 관한 법률', '파견근로자 보호 등에 관한 법률', '남녀고용평등과 일ㆍ가정 양립 지원에 관한 법률'
$precMeta = [ordered]@{}
foreach ($jo in $joList) {
  $list = Get-All @{ target = 'prec'; JO = $jo; curt = '대법원' } { param($r) [pscustomobject]@{ items = $r.PrecSearch.prec; total = $r.PrecSearch.totalCnt } } 5000
  foreach ($p in $list) {
    if (-not $precMeta.Contains($p.판례일련번호)) { $precMeta[$p.판례일련번호] = [pscustomobject]@{ id = $p.판례일련번호; 사건번호 = $p.사건번호; 선고일자 = $p.선고일자; 사건명 = $p.사건명; 참조법령 = $jo } }
  }
  Write-Host "판례 목록 [$jo]: $($list.Count)건 -> 누적 $($precMeta.Count)건"
}
$jobs = foreach ($id in $precMeta.Keys) {
  $f = Join-Path $data "prec\$id.json"
  if (-not (Test-Path $f)) { [pscustomobject]@{ url = "http://www.law.go.kr/DRF/lawService.do?OC=$OC&target=prec&ID=$id&type=JSON"; file = $f } }
}
Download-Many $jobs '판례 본문'
Save-Text (Join-Path $data 'prec\index.json') ($precMeta.Values | ConvertTo-Json -Depth 4)

# 4. 고용노동부 법령해석: 쟁점 키워드별로 끝까지
$expcMeta = [ordered]@{}
$list = Get-All @{ target = 'moelCgmExpc' } { param($r) [pscustomobject]@{ items = $r.CgmExpc.cgmExpc; total = $r.CgmExpc.totalCnt } } 20000
foreach ($e in $list) {
  if (-not $expcMeta.Contains($e.법령해석일련번호)) { $expcMeta[$e.법령해석일련번호] = [pscustomobject]@{ id = $e.법령해석일련번호; 안건번호 = $e.안건번호; 안건명 = $e.안건명 } }
}
Write-Host "해석 목록: $($list.Count)건 -> 고유 $($expcMeta.Count)건"
# 근로기준법 관련 쟁점만 받는다 (안건명 기준). 전체는 고용보험 등 다른 쟁점이 섞여 있다.
$rel = '근로기준법|해고|임금|근로시간|휴게|연차|휴가|퇴직|휴업|휴일|연장|야간|최저임금|산업재해|산재|재해|근로계약|징계|취업규칙|괴롭힘|상시|근로자|해직|체불|수당|평균|통상|사용자|사업장|가산|교대|수습|해약'
foreach ($id in @($expcMeta.Keys)) { if ($expcMeta[$id].안건명 -notmatch $rel) { $expcMeta.Remove($id) } }
Write-Host "관련 쟁점으로 거른 해석: $($expcMeta.Count)건"
$jobs = foreach ($id in $expcMeta.Keys) {
  $f = Join-Path $data "expc\$id.json"
  if (-not (Test-Path $f)) { [pscustomobject]@{ url = "http://www.law.go.kr/DRF/lawService.do?OC=$OC&target=moelCgmExpc&ID=$id&type=JSON"; file = $f } }
}
Download-Many $jobs '해석 본문'
Save-Text (Join-Path $data 'expc\index.json') ($expcMeta.Values | ConvertTo-Json -Depth 4)
Write-Host "완료: 판례 $($precMeta.Count)건, 해석 $($expcMeta.Count)건"
