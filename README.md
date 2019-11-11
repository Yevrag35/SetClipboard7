# Set-Clipboard v7

[![version](https://img.shields.io/powershellgallery/v/SetClipboard.svg)](https://www.powershellgallery.com/packages/SetClipboard)
[![downloads](https://img.shields.io/powershellgallery/dt/SetClipboard.svg?label=downloads)](https://www.powershellgallery.com/stats/packages/SetClipboard?groupby=Version)

This module brings the original functionality of the `Set-Clipboard` cmdlet to PowerShell 7.

[Set-Clipboard - MicrosoftDocs](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/set-clipboard)

When PowerShell 7 was first announced, one of the items brought up was that it runs on [.NET Core 3.0](https://devblogs.microsoft.com/powershell/powershell-7-road-map/), which brings back a lot of native Windows APIs including `System.Windows.Forms`.

The `Set-Clipboard` relies on the Forms assembly to invoke the `[System.Windows.Forms.Clipboard]` class, which is now back starting with PowerShell version 7.

I'm not sure if the PowerShell devs have it on their agenda to re-incorporate the cmdlet again, but until they do (or don't), this module can bring back the original functionality.  Although, I have not tested the `-AsHtml` switch yet.

---

## Installation

You can install it from the [PSGallery](https://www.powershellgallery.com) like you normally would a module.  It is suggested that you include the `-AllowClobber` parameter though, as PowerShell 7 can still detect that `Set-Clipboard` is a valid command for v5.1 and lower and can error saying the "command already exists".

### To install 'per-user'

``` powershell
Install-Module "SetClipboard" -AllowClobber
```

### To install for all users

``` powershell
Install-Module "SetClipboard" -Scope AllUsers -AllowClobber
```

---

## Examples

You can copy strings...

``` powershell
Set-Clipboard "whatev"
```

You can copy actual applications (not just their paths)...

``` powershell
Set-Clipboard -Path "C:\program.exe"
```

And you can invoke through the pipeline.

``` powershell
Get-Content "C:\log.txt" | Set-Clipboard
```
