function Invoke-Caesar {
    <#
    .SYNOPSIS
        An implementation of the Caesar shift cipher algorithm.

        Author: Alexandre Teyar (@aress31)
        License: BSD 3-Clause
        Required Dependencies: Windows Subsystem for Linux (WSL)
        Optional Dependencies: None

    .DESCRIPTION
        An implementation of the Caesar shift cipher algorithm.

    .PARAMETER Operation
        Specifies the operation to perform, i.e., encode or decode.

    .PARAMETER Key
        Specifies the key to use.

    .PARAMETER Value
        Specifies the value (hex-formatted) to encode or decode.
  
    .EXAMPLE
        PS C:\> cat payload.hex | Invoke-Caesar -Operation encode -Key 3

        Encode payload.hex using the Caesar shift cipher with a Key value of 3.

        PS C:\> cat payload.enc.hex | Invoke-Caesar -Operation decode -Key 3

        Decode payload.enc.hex using the Caesar shift cipher with a Key value of 3.
    #>
    
    [CmdletBinding(PositionalBinding = $False)]
    Param
    (
        [ValidateSet("decode", "encode")]
        [string]
        $Operation = "encode",

        [ValidateRange(0, 255)]
        [int]
        $Key = 1,

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
            $TransformedBytes += (($Byte + $Key) -band 255)
        }
        else {
            $TransformedBytes += (($Byte - $Key) -band 255)
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
    $HexArr = $Value -Split '(.{2})' -ne '' 
    [byte[]] $Bytes = @()

    foreach ($Hex in $HexArr) {
        $Bytes += [System.Convert]::ToByte($Hex, 16)
    }

    return $Bytes
}