# 링크 미리보기 이미지(og-image.png)와 파비콘(favicon.png)을 브라우저로 그려 scripts\assets 에 저장한다.
# 구글 폰트(Black Han Sans, Noto Sans KR)를 불러오므로 인터넷 연결과 브라우저가 필요하다.
# 문구나 숫자를 바꾸려면 scripts\og-design.html 을 고친 뒤 이 스크립트를 실행한다.
Add-Type -AssemblyName System.Web
$out = Join-Path $PSScriptRoot 'assets'
New-Item -ItemType Directory -Force $out | Out-Null
$design = Join-Path $PSScriptRoot 'og-design.html'
$port = 8767

$l = New-Object Net.HttpListener; $l.Prefixes.Add("http://localhost:$port/"); $l.Start()
$need = @{ 'og-image.png' = "http://localhost:$port/?v=a&save=og-image.png"; 'favicon.png' = "http://localhost:$port/?fav=1&save=favicon.png" }
$done = @{}
foreach ($u in $need.Values) { Start-Process $u }
Write-Host '브라우저에서 그림을 그리는 중입니다. 완료되면 자동으로 종료됩니다.'
$deadline = (Get-Date).AddMinutes(2)
while ($l.IsListening -and $done.Count -lt $need.Count -and (Get-Date) -lt $deadline) {
  $async = $l.BeginGetContext($null, $null)
  if (-not $async.AsyncWaitHandle.WaitOne(1000)) { continue }
  $c = $l.EndGetContext($async); $req = $c.Request; $res = $c.Response
  if ($req.HttpMethod -eq 'POST' -and $req.Url.AbsolutePath -eq '/save') {
    $name = [Web.HttpUtility]::ParseQueryString($req.Url.Query)['name'] -replace '[^\w\.\-]', ''
    $ms = New-Object IO.MemoryStream; $req.InputStream.CopyTo($ms)
    [IO.File]::WriteAllBytes((Join-Path $out $name), $ms.ToArray()); $done[$name] = 1; $res.StatusCode = 200
    Write-Host "저장: $name"
  } else {
    $b = [IO.File]::ReadAllBytes($design); $res.ContentType = 'text/html; charset=utf-8'; $res.OutputStream.Write($b, 0, $b.Length)
  }
  $res.Close()
}
$l.Stop()
Get-ChildItem $out | ForEach-Object { '{0}  {1:N0} KB' -f $_.Name, ($_.Length / 1KB) }
