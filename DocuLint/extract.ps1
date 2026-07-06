$projectRoot = $PSScriptRoot
$source = Join-Path $projectRoot "DocuLint\Ribbon\Ribbon1.cs"
$targetDir = Join-Path $projectRoot "DocuLint\Features\TablesAndFigures"
$target = "$targetDir\TablesAndFiguresRibbonActions.cs"
$csproj = Join-Path $projectRoot "DocuLint\DocuLint.csproj"

mkdir $targetDir -Force

$lines = Get-Content $source
$usings = $lines | Where-Object { $_ -match "^using " }

$startIndex = -1
for ($i=0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "private void button18_Click") {
        $startIndex = $i
        break
    }
}

$endIndex = $lines.Count - 3

$targetLines = @()
$targetLines += $usings
$targetLines += ""
$targetLines += "namespace DocuLint"
$targetLines += "{"
$targetLines += "    public partial class Ribbon1"
$targetLines += "    {"

for ($i = $startIndex; $i -le $endIndex; $i++) {
    $targetLines += $lines[$i]
}

$targetLines += "    }"
$targetLines += "}"

$targetLines | Out-File -FilePath $target -Encoding UTF8

$newSourceLines = @()
for ($i = 0; $i -lt $startIndex; $i++) {
    $newSourceLines += $lines[$i]
}
$newSourceLines += "    }"
$newSourceLines += "}"

$newSourceLines | Out-File -FilePath $source -Encoding UTF8

$csprojLines = Get-Content $csproj
$newCsprojLines = @()
$inserted = $false
for ($i = 0; $i -lt $csprojLines.Count; $i++) {
    if (-not $inserted -and $csprojLines[$i] -match "Features\\QuickTools\\QuickToolsRibbonActions.cs") {
        while ($csprojLines[$i] -notmatch "</Compile>") {
            $newCsprojLines += $csprojLines[$i]
            $i++
        }
        $newCsprojLines += $csprojLines[$i]
        $newCsprojLines += "    <Compile Include=`"Features\TablesAndFigures\TablesAndFiguresRibbonActions.cs`">"
        $newCsprojLines += "      <SubType>Component</SubType>"
        $newCsprojLines += "    </Compile>"
        $inserted = $true
    }
    else {
        $newCsprojLines += $csprojLines[$i]
    }
}

$newCsprojLines | Out-File -FilePath $csproj -Encoding UTF8
