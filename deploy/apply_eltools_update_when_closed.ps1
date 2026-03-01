$ErrorActionPreference = 'Stop'
$src = 'C:\Users\marat\ElTool\bin\Release\net8.0-windows'
$dst = 'C:\Users\marat\AppData\Roaming\Autodesk\ApplicationPlugins\ElTools.bundle\Contents\Win64'
$files = @('ElTools.dll','ElTools.deps.json','CommunityToolkit.Mvvm.dll','Newtonsoft.Json.dll','System.Management.dll','PanelLayoutMap.json')
for ($i=0; $i -lt 1800; $i++) {
    $acad = Get-Process acad -ErrorAction SilentlyContinue
    if ($null -eq $acad) { break }
    Start-Sleep -Seconds 1
}
foreach ($f in $files) {
    Copy-Item -Path (Join-Path $src $f) -Destination (Join-Path $dst $f) -Force
}
$marker = 'C:\Users\marat\ElTool\deploy\last_bundle_update.txt'
"UPDATED 1.5.4 $(Get-Date -Format o)" | Set-Content -Path $marker -Encoding utf8
