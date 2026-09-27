$ErrorActionPreference = "Stop"
$pptx = "D:\10_Workspace\translation-research-202610\问卷调查法在翻译研究中的正反例分析.pptx"
$out = "D:\10_Workspace\translation-research-202610\.pptbuild\qa"
if (Test-Path $out) { Remove-Item -Recurse -Force $out }
New-Item -ItemType Directory -Force -Path $out | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($pptx, $true, $false, $false)
  $pres.SaveAs($out, 18)   # 18 = ppSaveAsPNG, saves each slide as PNG into folder
  $pres.Close()
} finally {
  $ppt.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
Get-ChildItem $out | Select-Object Name, Length
