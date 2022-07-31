@{
    ModuleToProcess   = 'Vulcan.psm1'
    ModuleVersion     = '0.1.0.0'
    GUID              = '8443x493-q841-40d2-927a-0f09j228647x'
    Author            = 'Alexandre Teyar'
    Copyright         = 'BSD 3-Clause'
    Description       = 'A PowerShell script that simplifies life and therefore... phishing.'
    PowerShellVersion = '5.0'
    
    FunctionsToExport = @(
        'Invoke-Caesar',
        'Invoke-XOR',
        'Invoke-Vulcan'
    )
    
    ModuleList        = @( 
        @{ModuleName = 'Encoders'; ModuleVersion = '0.1.0.0'; GUID = '7cf9de61-2bwc-46b4-a397-9d7cf3a8e66b' },
        @{ModuleName = 'Utils'; ModuleVersion = '0.1.0.0'; GUID = 'a8a6280b-x694-4aa4-b28d-646afa66733c' }
    )
    
    PrivateData       = @{
        PSData = @{
            LicenseUri = 'https://opensource.org/licenses/BSD-3-Clause'
            ProjectUri = 'https://github.com/aress31/vulcan'
            Tags       = @(
                'security',
                'pentesting',
                'red team',
                'offense'
            )
        }
    }
}