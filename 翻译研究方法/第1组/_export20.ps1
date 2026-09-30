$ErrorActionPreference = "Stop"
$pptx = "D:\10_Workspace\translation-research-202610\翻译研究方法\第1组\问卷调查法正反例_v20.pptx"
$out = "D:\10_Workspace\translation-research-202610\翻译研究方法\第1组\_qa20"
if (Test-Path $out) { Remove-Item -Recurse -Force $out }
New-Item -ItemType Directory -Force -Path $out | Out-Null
$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($pptx, $true, $false, $false)
  $pres.SaveAs($out, 18)
  $pres.Close()
} finally {
  $ppt.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
"导出页数: " + (Get-ChildItem $out).Count
