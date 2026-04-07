param(
    [string]$Mode
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -f net8.0
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    if ([string]::IsNullOrWhiteSpace($Mode)) {
        dotnet run --project tests\Aspose.Cells_FOSS.CompareOpenXml\Aspose.Cells_FOSS.CompareOpenXml.csproj
    }
    else {
        dotnet run --project tests\Aspose.Cells_FOSS.CompareOpenXml\Aspose.Cells_FOSS.CompareOpenXml.csproj -- $Mode
    }

    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
    Pop-Location
}
