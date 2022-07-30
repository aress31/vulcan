function Invoke-Ceasar {
    <#
    .SYNOPSIS
        An implementation of the Caesar Shift cipher algorithm.

        Author: Alexandre Teyar (@aress31)
        License: BSD 3-Clause
        Required Dependencies: Windows Subsystem for Linux (WSL)
        Optional Dependencies: None

    .DESCRIPTION
        An implementation of the Caesar Shift cipher algorithm.

    .PARAMETER Operation
        Specifies the operation to perform, i.e., encode or decode.

    .PARAMETER Shift
        Specifies the shift to use.

    .PARAMETER Value
        Specifies the value (hex-formatted) to encode or decode.
  
    .EXAMPLE
        PS C:\> cat payload.hex | Invoke-Caesar -Operation encode -Shift 3

        Encode payload.hex using the Caesar Shift cipher with a shift value of 3.

        PS C:\> cat payload.enc.hex | Invoke-Caesar -Operation decode -Shift 3

        Decode payload.enc.hex using the Caesar Shift cipher with a shift value of 3.
    #>
    
    [CmdletBinding(PositionalBinding = $False)]
    [OutputType([String])]
    Param
    (
        [ValidateSet("decode", "encode")]
        [string]
        $Operation = "encode",

        [ValidateRange(0, 255)]
        [int]
        $Shift = 1,

        [Parameter(Mandatory, ValueFromPipeline)]
        [string]
        $Value
    )

    Write-Debug "Hex: $Value"

    $Bytes = Convert-HexToBytes -Value $Value
    [byte[]] $TransformedBytes = @()

    Write-Debug "Bytes: $Bytes"

    foreach ($Byte in $Bytes) {
        if ($Operation -eq "encode") {
            $TransformedBytes += (($Byte + $Shift) -band 255)
        }
        else {
            $TransformedBytes += (($Byte - $Shift) -band 255)
        }
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