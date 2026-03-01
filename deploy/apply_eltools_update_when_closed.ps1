$ErrorActionPreference = 'Stop'
$src = 'C:\Users\marat\ElTool\bin\Release\net8.0-windows'
$dst = 'C:\Users\marat\AppData\Roaming\Autodesk\ApplicationPlugins\ElTools.bundle\Contents\Win64'
$files = @(
    'ElTools.dll',
    'ElTools.deps.json',
    'ElTools.pdb',
    'PanelLayoutMap.json',
    'CommunityToolkit.Mvvm.dll',
    'Newtonsoft.Json.dll',
    'System.Management.dll',
    'ClosedXML.dll',
    'ClosedXML.Parser.dll',
    'DocumentFormat.OpenXml.dll',
    'DocumentFormat.OpenXml.Framework.dll',
    'ExcelNumberFormat.dll',
    'RBush.dll',
    'SixLabors.Fonts.dll',
    'System.IO.Packaging.dll'
)
if (-not (Test-Path $dst)) {
    New-Item -ItemType Directory -Path $dst -Force | Out-Null
}
for ($i = 0; $i -lt 1800; $i++) {
    $acad = Get-Process acad -ErrorAction SilentlyContinue
    if ($null -eq $acad) { break }
    Start-Sleep -Seconds 1
}
foreach ($f in $files) {
    Copy-Item -Path (Join-Path $src $f) -Destination (Join-Path $dst $f) -Force
}
$marker = 'C:\Users\marat\ElTool\deploy\last_bundle_update.txt'
"UPDATED 1.5.5 $(Get-Date -Format o)" | Set-Content -Path $marker -Encoding utf8
