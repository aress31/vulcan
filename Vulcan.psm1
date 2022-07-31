Get-ChildItem -Path $PSScriptRoot -File -Filter *.ps1 | `
    ForEach-Object { Import-Module $_.FullName }

# Get-ChildItem -Path $PSScriptRoot -Directory | `
#     Where-Object { !(@('assets') -contains $_.Name) } | `
#     ForEach-Object { Import-Module $_.FullName }