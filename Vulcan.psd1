@{
    RootModule        = 'Vulcan'
    ModuleVersion     = '0.0.1'
    GUID              = 'fbe14f37-d148-495e-8b55-febf6579860e'
    Author            = 'Alexandre Teyar'
    CompanyName       = 'Aegis Cyber Ltd.'
    Copyright         = '(c) 2022 Alexandre Teyar. All rights reserved.'
    Description       = 'A PowerShell script that simplifies life and therefore... phishing.'
    PowerShellVersion = '5.0'
     
    NestedModules     = @( 
        @{
            ModuleName    = 'Encoder'
            ModuleVersion = '0.0.1'
            GUID          = 'f4768a29-203a-443a-9cf6-d54bb241cf9c'
        }
    )

    FunctionsToExport = @(
        'Invoke-Caesar',
        'Invoke-Vulcan',
        'Invoke-XOR'
    )

    ModuleList        = @(
        @{
            ModuleName    = 'Encoder'
            ModuleVersion = '0.0.1'
            GUID          = 'f4768a29-203a-443a-9cf6-d54bb241cf9c'
        }
    )
    
    FileList          = @(
        'Vulcan.ps1'
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