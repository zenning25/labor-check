# data\ 의 수집 결과를 요약해 템플릿에 내장하고 근로기준법_판별기_로컬.html 을 만든다.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$data = Join-Path $root 'data'
$utf8 = New-Object Text.UTF8Encoding($false)

function Read-Json($path) { try { ConvertFrom-Json ([IO.File]::ReadAllText($path, $utf8)) } catch { Write-Host "읽기 실패, 건너뜀: $path"; $null } }
function Clean($s, $max) {
  if ($null -eq $s) { return '' }
  $t = [string]$s
  $t = $t -replace '<br\s*/?>', "`n" -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&'
  $t = ($t -replace '[ \t ]+', ' ' -replace '(\s*\n\s*)+', "`n").Trim()
  if ($max -and $t.Length -gt $max) { $t = $t.Substring(0, $max) + '…' }
  return $t
}
function Flatten($node) {
  if ($null -eq $node -or $node -is [string]) { return }
  if ($node -is [array]) { foreach ($n in $node) { Flatten $n }; return }
  foreach ($p in $node.PSObject.Properties) {
    if ($p.Name -match '내용$') { foreach ($v in @($p.Value)) { if ($v -is [string] -and $v.Trim()) { $v.Trim() } } }
    elseif ($p.Value -isnot [string]) { Flatten $p.Value }
  }
}

# 법령
$laws = [ordered]@{}
foreach ($f in @{ '근로기준법' = '근로기준법.json'; '시행령' = '근로기준법_시행령.json'; '시행규칙' = '근로기준법_시행규칙.json' }.GetEnumerator()) {
  $j = Read-Json (Join-Path $data "law\$($f.Value)")
  $arts = @()
  foreach ($a in @($j.법령.조문.조문단위)) {
    if ($a.조문여부 -ne '조문') { continue }
    $key = [string]$a.조문번호
    if ($a.조문가지번호) { $key += "의$($a.조문가지번호)" }
    $lines = @(([string]$a.조문내용).Trim()) + @(Flatten $a.항)
    $arts += [pscustomobject]@{ key = $key; title = [string]$a.조문제목; text = (Clean ($lines -join "`n") 0) }
  }
  $laws[$f.Key] = $arts
  $info = $j.법령.기본정보
  Write-Host "$($f.Key): 조문 $($arts.Count)개 (시행 $($info.시행일자))"
  $laws["$($f.Key)_시행일"] = [string]$info.시행일자
}

# 시행령 별표 1
$sl = Read-Json (Join-Path $data 'law\근로기준법_시행령.json')
$byl1 = ''
foreach ($u in @($sl.법령.별표.별표단위)) { if ($u.별표번호 -eq '0001') { $byl1 = (($u.별표내용 | Out-String -Width 4096) -replace '[ \t]+(\r?\n)', '$1').Trim(); break } }
if (-not $byl1) { throw '별표 1을 찾지 못했습니다.' }

# 다른 노동 법령 조문 + 적용 메모 (collect_laws.ps1 이 만든 data\laws)
$xlaws = [ordered]@{}
$xscope = $null
$xdir = Join-Path $data 'laws'
if (Test-Path (Join-Path $xdir 'index.json')) {
  $xscope = ConvertFrom-Json ([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'law_scope.json'), $utf8))
  foreach ($i in (@(Read-Json (Join-Path $xdir 'index.json')) | ForEach-Object { $_ })) {
    $j = Read-Json (Join-Path $xdir $i.파일)
    if (-not $j) { continue }
    $arts = New-Object Collections.Generic.List[object]
    foreach ($a in @($j.법령.조문.조문단위)) {
      if ($a.조문여부 -ne '조문') { continue }
      $key = [string]$a.조문번호
      if ($a.조문가지번호) { $key += "의$($a.조문가지번호)" }
      $lines = @(([string]$a.조문내용).Trim()) + @(Flatten $a.항)
      $arts.Add([pscustomobject]@{ k = $key; t = [string]$a.조문제목; x = (Clean ($lines -join "`n") 0) })
    }
    foreach ($u in @($j.법령.별표.별표단위)) {
      $bt = ($u.별표내용 | Out-String -Width 4096)
      if ($bt -match '\[별표 1\]') { $arts.Add([pscustomobject]@{ k = '별표1'; t = '별표 1'; x = (($bt -replace '[ \t]+(\r?\n)', '$1').Trim()) }); break }
    }
    $xlaws[$i.법령명] = [pscustomobject]@{ since = [string]$i.시행일자; arts = $arts.ToArray() }
  }
  Write-Host "다른 법령: $($xlaws.Count)개"
}

# 판례 (참조법령 라벨은 collect.ps1 이 만든 index.json 에서 가져온다)
$joMap = @{}
$idxFile = Join-Path $data 'prec\index.json'
if (Test-Path $idxFile) { foreach ($i in (@(Read-Json $idxFile) | ForEach-Object { $_ })) { $joMap[[string]$i.id] = [string]$i.참조법령 } }
$prec = New-Object Collections.Generic.List[object]
foreach ($f in Get-ChildItem (Join-Path $data 'prec') -File | Where-Object Name -ne 'index.json') {
  $p = (Read-Json $f.FullName).PrecService
  if (-not $p) { continue }
  $prec.Add([pscustomobject]@{
    id = $f.BaseName; no = [string]$p.사건번호; date = [string]$p.선고일자; type = [string]$p.사건종류명; name = [string]$p.사건명
    jo = [string]$joMap[$f.BaseName]; issue = (Clean $p.판시사항 700); gist = (Clean $p.판결요지 1400); refs = (Clean $p.참조조문 400)
  })
}
# 해석
$expc = New-Object Collections.Generic.List[object]
foreach ($f in Get-ChildItem (Join-Path $data 'expc') -File | Where-Object Name -ne 'index.json') {
  $e = (Read-Json $f.FullName).CgmExpcService
  if (-not $e) { continue }
  $expc.Add([pscustomobject]@{
    id = $f.BaseName; no = [string]$e.안건번호; date = [string]$e.해석일자; title = [string]$e.안건명
    q = (Clean $e.질의요지 450); a = (Clean $e.회답 750)
  })
}
Write-Host "판례 $($prec.Count)건, 해석 $($expc.Count)건"
$built = Get-Date -Format 'yyyy-MM-dd'
$tpl = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'template.html'), $utf8)
function To-Json($o) { (ConvertTo-Json -InputObject $o -Depth 8 -Compress) -replace '</', '<\/' }

# 1) 호스팅용: docs\index.html + docs\data\*.json (전체 데이터)
$docs = Join-Path $root 'docs'
New-Item -ItemType Directory -Force (Join-Path $docs 'data') | Out-Null
[IO.File]::WriteAllText((Join-Path $docs 'index.html'), $tpl.Replace('/*__DATA__*/null', 'null'), $utf8)
[IO.File]::WriteAllText((Join-Path $docs 'data\law.json'), (To-Json ([ordered]@{ built = $built; laws = $laws; byl1 = $byl1; xlaws = $xlaws; xscope = $xscope })), $utf8)
[IO.File]::WriteAllText((Join-Path $docs 'data\prec.json'), (To-Json $prec.ToArray()), $utf8)
[IO.File]::WriteAllText((Join-Path $docs 'data\expc.json'), (To-Json $expc.ToArray()), $utf8)
[IO.File]::WriteAllText((Join-Path $docs '.nojekyll'), '', $utf8)
Write-Host ("저장: docs\ ({0:N1} MB)" -f ((Get-ChildItem $docs -Recurse -File | Measure-Object Length -Sum).Sum / 1MB))

# 2) 로컬 단일 파일: 크기를 줄이기 위해 요지를 더 줄이고 근로기준법과 관련 낮은 해석은 뺀다
$rel = '근로기준법|해고|임금|근로시간|휴게|연차|휴가|퇴직|휴업|휴일|연장|야간|최저임금|산업재해|산재|근로계약|징계|취업규칙|괴롭힘|상시|근로자'
$prec2 = $prec | ForEach-Object { [pscustomobject]@{ id = $_.id; no = $_.no; date = $_.date; type = $_.type; name = $_.name; jo = $_.jo
  issue = $(if ($_.issue.Length -gt 450) { $_.issue.Substring(0, 450) + '…' } else { $_.issue })
  gist = $(if ($_.gist.Length -gt 800) { $_.gist.Substring(0, 800) + '…' } else { $_.gist })
  refs = $(if ($_.refs.Length -gt 250) { $_.refs.Substring(0, 250) } else { $_.refs }) } }
$expc2 = $expc | Where-Object { ($_.title + $_.q) -match $rel } | ForEach-Object { [pscustomobject]@{ id = $_.id; no = $_.no; date = $_.date; title = $_.title
  q = $(if ($_.q.Length -gt 300) { $_.q.Substring(0, 300) + '…' } else { $_.q })
  a = $(if ($_.a.Length -gt 500) { $_.a.Substring(0, 500) + '…' } else { $_.a }) } }
Write-Host "로컬 파일에 넣는 양: 판례 $(@($prec2).Count)건, 해석 $(@($expc2).Count)건"
$payload = [ordered]@{ built = $built; laws = $laws; byl1 = $byl1; xlaws = $xlaws; xscope = $xscope; prec = @($prec2); expc = @($expc2) }
$html = $tpl.Replace('/*__DATA__*/null', (To-Json $payload))
$dest = Join-Path $root '근로기준법_판별기_로컬.html'
[IO.File]::WriteAllText($dest, $html, $utf8)
Write-Host ("저장: $dest ({0:N1} MB)" -f ((Get-Item $dest).Length / 1MB))

# 3) claude.ai Artifact 게시용: html/head/body 래퍼 없이 title, meta, style, body 내용만 담는다.
$style = [regex]::Match($html, '(?s)<style>.*?</style>').Value
$body = [regex]::Match($html, '(?s)<body>(.*)</body>').Groups[1].Value
$pub = "<title>근로기준법 판별기</title>`n<meta name=`"claude:viewport-fit-sides`" content=`"cover`">`n$style`n$body"
[IO.File]::WriteAllText((Join-Path $root '근로기준법_판별기_게시용.html'), $pub, $utf8)
Write-Host '저장: 근로기준법_판별기_게시용.html'