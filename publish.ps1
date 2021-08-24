if (Test-Path '.\published') {
    Remove-Item '.\published'
}

& dotnet publish YearGraphs\YearGraphs.fsproj -c Release -r win-x64 --self-contained true -o .\published -p:PublishSingleFile=true -p:PublishTrimmed=true -p:PublishReadyToRun=true