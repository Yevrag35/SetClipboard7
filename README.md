# Set-Clipboard v7

This module brings the original functionality of the `Set-Clipboard` cmdlet to PowerShell 7.

[Set-Clipboard - MicrosoftDocs](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/set-clipboard)

When PowerShell 7 was first announced, one of the items brought up was that it runs on [.NET Core 3.0](https://devblogs.microsoft.com/powershell/powershell-7-road-map/), which brings back a lot native Windows API including `System.Windows.Forms`.

The `Set-Clipboard` relies on the Forms assembly to invoke `[System.Windows.Forms.Clipboard]` class, which is now back starting with PowerShell version 7.

I'm not sure if the PowerShell devs have it on their agenda to re-incorporate the cmdlet again, but until they do (or don't), this module can bring back the original functionality.  Although, I have not tested the `-AsHtml` switch yet.

## Examples

``` powershell
Set-Clipboard "whatev"
```

``` powershell
Set-Clipboard -Path "C:\program.exe"
```

``` powershell
Get-Content "C:\log.txt" | Set-Clipboard
```
