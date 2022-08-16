Get-ChildItem -Path $PSScriptRoot -File -Filter *.ps1 -Recurse | `
    ForEach-Object { Import-Module $_.FullName }