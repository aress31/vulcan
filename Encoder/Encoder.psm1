Get-ChildItem -Path $PSScriptRoot -Filter *.ps1 -File | `
ForEach-Object { . $_.FullName }