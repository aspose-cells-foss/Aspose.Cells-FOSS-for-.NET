$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -f net8.0
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $projects = @(
        'tests\Aspose.Cells_FOSS.UnitTests\Aspose.Cells_FOSS.UnitTests.csproj',
        'tests\Aspose.Cells_FOSS.GoldenTests\Aspose.Cells_FOSS.GoldenTests.csproj',
        'tests\Aspose.Cells_FOSS.MalformedTests\Aspose.Cells_FOSS.MalformedTests.csproj',
        'tests\Aspose.Cells_FOSS.CompatibilityTests\Aspose.Cells_FOSS.CompatibilityTests.csproj'
    )

    foreach ($project in $projects) {
        dotnet run --project $project
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    }
}
finally {
    Pop-Location
}
