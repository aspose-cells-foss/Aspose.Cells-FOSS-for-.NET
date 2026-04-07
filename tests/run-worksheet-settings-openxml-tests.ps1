$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -f net8.0
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    dotnet run --project tests\Aspose.Cells_FOSS.WorksheetSettingsOpenXmlTests\Aspose.Cells_FOSS.WorksheetSettingsOpenXmlTests.csproj
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
    Pop-Location
}
