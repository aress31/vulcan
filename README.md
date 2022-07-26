# vulcan

[![Language](https://img.shields.io/badge/Lang-PowerShell-blue.svg)](https://docs.microsoft.com/en-gb/powershell/)
[![License](https://img.shields.io/badge/License-BSD%203-red.svg)](https://opensource.org/licenses/BSD-3-Clause)

## A collection of `PowerShell` scripts to simplify everyday phishing. 🎣

For now, a simple `PowerShell` script to automate the entire process of generating consistent and efficient macro-enabled `Word` documents that actually manage to evade - *at the time of writing* - most AVs, including `Windows Defender`. This script is a wrapper around [MSFVenom](https://www.offensive-security.com/metasploit-unleashed/msfvenom/) - made possible thanks to [Windows Subsystem for Linux](https://docs.microsoft.com/en-us/windows/wsl/install) and [BadAssMacros](https://github.com/Inf0secRabbit/BadAssMacros).

More features might follow, but this repository mainly serves as a place for the community to build and improve upon the existing. 🐺

## ⚠️ Do not submit the produced files to VirusTotal.

## Requirements

- [BadAssMacros](https://github.com/Inf0secRabbit/BadAssMacros)
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
    Invoke-Vulcan -BadAssMacros $BadAssMacros -Payload $Payload -PayloadOptions $PayloadOptions
    ```

> Although obvious, `$BadAssMacros`, `$Payload` and `$PayloadOptions` are placeholders for your own values... 🙄

## Usage

`Get-Help -Name Invoke-Vulcan` is your friend... Even better if you use `Get-Help -Name Invoke-Vulcan -Detailed`.

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

- [BadAssMacros](https://github.com/Inf0secRabbit/BadAssMacros)
- [MSFVenom](https://www.offensive-security.com/metasploit-unleashed/msfvenom/)

## License

`vulcan` is primarily distributed under the terms of the `BSD 3`.

See [LICENSE](./LICENSE) for details.