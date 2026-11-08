# PowerPoint 実機（COM）でスライドを JPEG に書き出し、実際に表示された行を lines.tsv に保存する。
# 使い方: powershell -ExecutionPolicy Bypass -File check_ppt.ps1 -Path <pptxのフルパス> -Out <出力フォルダ>
# 出力: slide01.jpg ...、lines.tsv（slide, shape, id, para, line, text）→ lint.py --lines で使う
# 図形は id で照合する（COM は既定のプレースホルダー名を英語で返し、ファイル内の名前と一致しないため）
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path $Path).Path
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
$rows = New-Object System.Collections.Generic.List[string]
$rows.Add("slide`tshape`tid`tpara`tline`ttext")

function Dump-Text($slideNo, $name, $id, $tf) {
  if (-not $tf.HasText) { return }
  $tr = $tf.TextRange
  for ($p = 1; $p -le $tr.Paragraphs().Count; $p++) {
    $para = $tr.Paragraphs($p)
    for ($l = 1; $l -le $para.Lines().Count; $l++) {
      $t = $para.Lines($l).Text -replace "[`r`n`v]", "" -replace "`t", " "
      $rows.Add("$slideNo`t$name`t$id`t$p`t$l`t$t")
    }
  }
}

function Walk($slideNo, $shapes, $prefix) {
  foreach ($sh in $shapes) {
    $name = "$prefix$($sh.Name)"
    if ($sh.Type -eq 6) { Walk $slideNo $sh.GroupItems "$name/"; continue }   # msoGroup
    if ($sh.HasTable) {
      $tb = $sh.Table
      for ($r = 1; $r -le $tb.Rows.Count; $r++) {
        for ($c = 1; $c -le $tb.Columns.Count; $c++) {
          # lint.py と合わせるため 0 始まりで出力
          $cell = "[{0},{1}]" -f ($r - 1), ($c - 1)
          Dump-Text $slideNo "$name$cell" "$($sh.Id)$cell" $tb.Cell($r, $c).Shape.TextFrame
        }
      }
      continue
    }
    if ($sh.HasTextFrame) { Dump-Text $slideNo $name "$($sh.Id)" $sh.TextFrame }
  }
}

$ppt = New-Object -ComObject PowerPoint.Application
# ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoFalse
$pres = $ppt.Presentations.Open($Path, -1, 0, 0)
try {
  $w = 1920
  $h = [int](1920 * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  foreach ($s in $pres.Slides) {
    $s.Export((Join-Path $Out ("slide{0:D2}.jpg" -f $s.SlideIndex)), "JPG", $w, $h)
    Walk $s.SlideIndex $s.Shapes ""
  }
} finally {
  $pres.Close()
  $ppt.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
[System.IO.File]::WriteAllLines((Join-Path $Out "lines.tsv"), $rows, (New-Object System.Text.UTF8Encoding $false))
Write-Host "書き出し完了: $Out"
