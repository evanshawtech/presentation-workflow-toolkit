param(
  [Parameter(Mandatory=$true)][string]$InputPptx,
  [Parameter(Mandatory=$true)][string]$OutDir,
  [int]$Width = 1600
)
$ErrorActionPreference = 'Stop'
$inputFull = (Resolve-Path -LiteralPath $InputPptx).Path
if ([IO.Path]::GetExtension($inputFull) -ne '.pptx') { throw 'Only .pptx accepted.' }
$outputFull = [IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $outputFull | Out-Null
$renderMutex = New-Object System.Threading.Mutex($false, 'Local\CodexPptToolkitRenderer')
try { $acquired = $renderMutex.WaitOne(120000) }
catch [System.Threading.AbandonedMutexException] { $acquired = $true }
if (-not $acquired) { $renderMutex.Dispose(); throw 'Another toolkit render is still using PowerPoint.' }
$alreadyRunning = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$application = $null
$presentation = $null
$originalSecurity = $null
try {
  $application = New-Object -ComObject PowerPoint.Application
  $originalSecurity = $application.AutomationSecurity
  $application.AutomationSecurity = 3
  for ($attempt=0; $attempt -lt 3; $attempt++) {
    try { $presentation = $application.Presentations.Open($inputFull, -1, 0, 0); break }
    catch { if ($attempt -eq 2) { throw }; Start-Sleep -Milliseconds 750 }
  }
  $height = [int][Math]::Round($Width * $presentation.PageSetup.SlideHeight / $presentation.PageSetup.SlideWidth)
  $count = $presentation.Slides.Count
  for ($i=1; $i -le $count; $i++) {
    $slide = $presentation.Slides.Item($i)
    $slide.Export((Join-Path $outputFull ('slide-{0:D3}.png' -f $i)), 'PNG', $Width, $height)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($slide)
  }
  $presentation.SaveAs((Join-Path $outputFull 'rendered.pdf'), 32)
  @{engine='Microsoft PowerPoint';opened=$true;slide_count=$count;width=$Width;height=$height;input=$inputFull} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputFull 'render.json') -Encoding utf8
  Write-Output "PowerPoint opened and rendered $count slides."
} finally {
  if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
  if ($application) {
    if ($null -ne $originalSecurity) { $application.AutomationSecurity = $originalSecurity }
    if (-not $alreadyRunning -and $application.Presentations.Count -eq 0) { $application.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  $renderMutex.ReleaseMutex()
  $renderMutex.Dispose()
}
