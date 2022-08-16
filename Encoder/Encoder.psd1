@{
    ModuleToProcess   = 'Vulcan.psm1'
    ModuleVersion     = '0.1.0.0'
    GUID              = '937gh27a3-q841-40d2-927a-0f09j228647x'
    Author            = 'Alexandre Teyar'
    Copyright         = 'BSD 3-Clause'
    Description       = 'A collection of encoders/decoders.'
    PowerShellVersion = '5.0'
    
    FunctionsToExport = @(
        'Invoke-Caesar',
        'Invoke-XOR'
    )
    
    FileList          = 'Caesar.ps1', 'XOR.ps1'
}