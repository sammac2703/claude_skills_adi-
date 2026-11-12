param([string]$In, [string]$OutDir, [int]$Width = 960)
New-Item -ItemType Directory -Force $OutDir | Out-Null
$app = New-Object -ComObject PowerPoint.Application
try {
    # ReadOnly, Untitled, WithWindow=false
    $pres = $app.Presentations.Open($In, -1, -1, 0)
    $h = [int]($Width * 9 / 16)
    $pres.Export($OutDir, "PNG", $Width, $h)
    $pres.Close()
} finally {
    $app.Quit()
}
Get-ChildItem $OutDir -Filter *.PNG | Measure-Object | ForEach-Object { "$($_.Count) slides exported" }
