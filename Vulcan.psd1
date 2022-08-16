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
        @{ModuleName = 'Encoders'; ModuleVersion = '0.1.0.0'; GUID = '7cf9de61-2bwc-46b4-a397-9d7cf3a8e66b' },
        @{ModuleName = 'Utils'; ModuleVersion = '0.1.0.0'; GUID = 'a8a6280b-x694-4aa4-b28d-646afa66733c' }
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