@{
    RootModule        = 'Encoder'
    ModuleVersion     = '0.0.1'
    GUID              = 'f4768a29-203a-443a-9cf6-d54bb241cf9c'
    Author            = 'Alexandre Teyar'
    Copyright         = '(c) 2022 Alexandre Teyar. All rights reserved.'
    Description       = 'A collection of PowerShell encoders/decoders.'
    PowerShellVersion = '5.0'

    FunctionsToExport = @(
        'Invoke-Caesar',
        'Invoke-XOR'
    )

    FileList          = @(
        'Caesar.ps1',
        'XOR.ps1'
    )

    PrivateData       = @{
        PSData = @{
            Tags                     = @(
                'offense',
                'pentesting',
                'red team',
                'security'
            )
            LicenseUri               = 'https://opensource.org/licenses/BSD-3-Clause'
            ProjectUri               = 'https://github.com/aress31/vulcan'
            RequireLicenseAcceptance = $true
        }
    }
}