function Invoke-XOR {
    <#
    .SYNOPSIS
        An implementation of the XOR algorithm.

        Author: Alexandre Teyar (@aress31)
        License: BSD 3-Clause
        Required Dependencies: Windows Subsystem for Linux (WSL)
        Optional Dependencies: None

    .DESCRIPTION
        An implementation of the XOR algorithm.

    .PARAMETER Key
        Specifies the key/secret to use.

    .PARAMETER Value
        Specifies the value (hex-formatted) to encode or decode.
  
    .EXAMPLE
        PS C:\> cat payload.hex | Invoke-XOR -Operation encode -Shift 3

        Encode payload.hex using the key of 3.

        PS C:\> cat payload.enc.hex | Invoke-Caesar -Operation decode -Shift 3

        Decode payload.enc.hex using the Caesar Shift cipher with a shift value of 3.
    #>
    
    [CmdletBinding(PositionalBinding = $False)]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory)]
        [string]
        $Key,

        [Parameter(Mandatory, ValueFromPipeline)]
        [string]
        $Value
    )

    Write-Debug "Hex: $Value"

    $Bytes = Convert-HexToBytes -Value $Value
    [byte[]] $TransformedBytes = @()

    Write-Debug "Bytes: $Bytes"

    for ($i = 0; $i -lt $Bytes.Count; $i++) {
        $TransformedBytes += $Bytes[$i] -bxor $Key[$i % $Key.Length]
    }
    
    Write-Debug "TransformedBytes: $TransformedBytes"
    $TransformedHex = Convert-BytesToHex -Value $TransformedBytes

    Write-Debug "TransformedHex: $TransformedHex"

    return $TransformedHex.ToLower()
}

function  Convert-BytesToHex($Value) {
    return [System.BitConverter]::ToString($Value).Replace('-', '')
}

function Convert-HexToBytes($Value) {
    $HexArray = $Value -Split '(.{2})' -ne '' 
    [byte[]]$Bytes = @()

    foreach ($Hex in $HexArray) {
        $Bytes += [System.Convert]::ToByte($Hex, 16)
    }

    return $Bytes
}