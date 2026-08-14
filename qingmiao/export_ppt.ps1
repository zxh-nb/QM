param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Source)) { throw "Source not found: $Source" }
$Source = (Resolve-Path $Source).Path
$OutDir = (New-Item -ItemType Directory -Force -Path $OutDir).FullName

$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1
try {
    $pres = $app.Presentations.Open($Source, $true, $false, $false)
    $count = $pres.Slides.Count
    Write-Output "Slides: $count"
    for ($i = 1; $i -le $count; $i++) {
        $outPath = Join-Path $OutDir ("slide-{0:D2}.png" -f $i)
        $pres.Slides.Item($i).Export($outPath, "PNG", $Width, $Height)
        Write-Output "Exported $outPath"
    }
    $pres.Close()
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
