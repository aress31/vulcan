Get-ChildItem $PSScriptRoot | `
    Where-Object { $_.PSIsContainer } | `
    ForEach-Object { Import-Module $_.FullName -DisableNameChecking }