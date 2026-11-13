# 링크 미리보기 이미지(og-image.png)와 파비콘(favicon.png)을 그린다. 결과는 scripts\assets 에 저장된다.
Add-Type -AssemblyName System.Drawing
$out = Join-Path $PSScriptRoot 'assets'
New-Item -ItemType Directory -Force $out | Out-Null

function RGB($hex, $a = 255) { $h = $hex.TrimStart('#'); [Drawing.Color]::FromArgb($a, [Convert]::ToInt32($h.Substring(0,2),16), [Convert]::ToInt32($h.Substring(2,2),16), [Convert]::ToInt32($h.Substring(4,2),16)) }
function RoundRect($x, $y, $w, $h, $r) {
  $p = New-Object Drawing.Drawing2D.GraphicsPath
  $d = 2 * $r
  $p.AddArc($x, $y, $d, $d, 180, 90); $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90); $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $p.CloseFigure(); return $p
}
function NewCanvas($w, $h) {
  $bmp = New-Object Drawing.Bitmap($w, $h)
  $g = [Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'; $g.InterpolationMode = 'HighQualityBicubic'
  return @($bmp, $g)
}
function Font($size, $bold = $true) { New-Object Drawing.Font('Malgun Gothic', $size, $(if ($bold) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }), [Drawing.GraphicsUnit]::Pixel) }

# ---- 미리보기 이미지 1200x630 ----
$bmp, $g = NewCanvas 1200 630
$g.Clear((RGB '#FFD43B'))
# 배경 장식: 연한 큰 원 두 개 (오른쪽 아래), 빨간 점
$g.FillEllipse((New-Object Drawing.SolidBrush((RGB '#FFE27A'))), 760, 250, 620, 620)
$g.FillEllipse((New-Object Drawing.SolidBrush((RGB '#FFDF6B'))), 940, 40, 360, 360)
$ink = New-Object Drawing.SolidBrush((RGB '#191F28'))
$sub = New-Object Drawing.SolidBrush((RGB '#4A3B00'))
$g.FillEllipse((New-Object Drawing.SolidBrush((RGB '#E5322D'))), 80, 78, 26, 26)
$g.DrawString('LABOR LAW CHECK', (Font 26), $sub, 120, 72)
$g.DrawString('근로기준법 판별기', (Font 112), $ink, 70, 150)
$g.DrawString('5인 미만 사업장 특례까지 한 번에 확인', (Font 46 $false), $ink, 80, 308)

# 정보 알약 3개
$fp = Font 27
$x = 80; $y = 430
foreach ($t in '5인 미만 특례 대조', '대법원 판례 2,612건', '고용노동부 해석 4,128건') {
  $w = [int]($g.MeasureString($t, $fp).Width) - 6 + 48
  $path = RoundRect $x $y $w 68 34
  $g.FillPath((New-Object Drawing.SolidBrush((RGB '#FFFFFF'))), $path)
  $g.DrawString($t, $fp, $ink, ($x + 24), ($y + 17))
  $x += $w + 16
}
"알약 오른쪽 끝: $($x - 16) (이미지 너비 1200, 오른쪽 여백 80이면 1120 이하)"
$g.DrawString('zenning25.github.io/labor-check', (Font 28 $false), $sub, 80, 548)
$bmp.Save((Join-Path $out 'og-image.png'), [Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()

# ---- 파비콘 128x128 (노란 둥근 사각형 + 판) ----
$bmp, $g = NewCanvas 128 128
$g.Clear([Drawing.Color]::Transparent)
$g.FillPath((New-Object Drawing.SolidBrush((RGB '#FFD43B'))), (RoundRect 0 0 128 128 30))
$g.FillEllipse((New-Object Drawing.SolidBrush((RGB '#E5322D'))), 20, 18, 16, 16)
$fmt = New-Object Drawing.StringFormat; $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
$g.DrawString('판', (Font 78), (New-Object Drawing.SolidBrush((RGB '#191F28'))), (New-Object Drawing.RectangleF(0, 10, 128, 120)), $fmt)
$bmp.Save((Join-Path $out 'favicon.png'), [Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Get-ChildItem $out | ForEach-Object { '{0}  {1:N0} KB' -f $_.Name, ($_.Length / 1KB) }
