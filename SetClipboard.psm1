#region BACKEND/SUPPORTING FUNCTIONS

Function Copy-FilesToClipboard()
{
    [CmdletBinding(PositionalBinding=$false)]
    param
    (
        [System.Collections.Generic.List[string]]$fileList,
        [bool]$append,
        [bool]$isLiteralPath
    )
    [int]$count = 0
    $source = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($append)
    {
        if (-not [System.Windows.Forms.Clipboard]::ContainsFileDropList())
        {
            Write-Verbose "No appendable content was found."
            $append = $false
        }
        else
        {
            [System.Collections.Generic.List[string]]$fileDropList = [System.Windows.Forms.Clipboard]::GetFileDropList()
            $source = [System.Collections.Generic.HashSet[string]]::new($fileDropList, [System.StringComparer]::OrdinalIgnoreCase)
            $count = $fileDropList.Count
        }
    }

    New-Variable -Name "providerInfo" -Option AllScope -Value $null
    for ($i = 0; $i -lt $fileList.Count; $i++)
    {
        $resolvedProviderPathFromPSPath = New-Object -TypeName 'System.Collections.ObjectModel.Collection[string]'
        try
        {
            if ($isLiteralPath)
            {
                [void]$resolvedProviderPathFromPSPath.Add($PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($fileList[$i]))
            }
            else
            {

                [void]$resolvedProviderPathFromPSPath.Add($PSCmdlet.SessionState.Path.GetResolvedProviderPathFromPSPath($fileList[$i], [ref] $providerInfo))
            }
        }
        catch [System.Management.Automation.ItemNotFoundException]
        {
            Write-Error -ErrorRecord $([System.Management.Automation.ErrorRecord]::new($_.Exception, "FailedToSetClipboard", [System.Management.Automation.ErrorCategory]::InvalidOperation, "Clipboard"))
        }
        foreach ($str2 in $resolvedProviderPathFromPSPath)
        {
            if (-not $source.Contains($str2))
            {
                [void]$source.Add($str2)
            }
        }
    }

    if ($source.Count -ne 0)
    {
        if (($source.Count - $count) -eq 1)
        {
            if ($append)
            {
                $msg = "Append single file '{0}' to the clipboard." -f $($source[$source.Count - 1])
            }
            else
            {
                $msg = "Set single file '{0}' to the clipboard." -f $source[0]
            }
        }
        elseif ($append)
        {
            $msg = "Append {0} files to the clipboard." -f ($source.Count - $count)
        }
        else
        {
            $msg = "Set {0} files to the clipboard." -f $source.Count
        }

        if ($PSCmdlet.ShouldProcess($msg, "Set-Clipboard"))
        {
            [void][System.Windows.Forms.Clipboard]::Clear()
            $filePaths = [System.Collections.Specialized.StringCollection]::new()
            [string[]]$strs = $resolvedProviderPathFromPSPath
            $filePaths.AddRange($strs)
            $filePaths
            [System.Windows.Forms.Clipboard]::SetFileDropList($filePaths)
        }
    }
}

Function Get-ByteCount([System.Text.StringBuilder]$sb, [int]$start = 0, [int]$end = -1)
{
    $chars = New-Object char[] 1
    [int]$num = 0
    $end = ($end -gt -1) ? $end : $sb.Length
    for ($i = $start; $i -lt $end; $i++)
    {
        $chars[0] = $sb[$i]
        $num += [System.Text.Encoding]::UTF8.GetByteCount($chars)
    }
    $num
}

Function Get-HtmlDataString([string]$html)
{
    $bsString = "Version:0.9\r\nStartHTML:<<<<<<<<1\r\nEndHTML:<<<<<<<<2\r\nStartFragment:<<<<<<<<3\r\nEndFragment:<<<<<<<<4\r\nStartSelection:<<<<<<<<3\r\nEndSelection:<<<<<<<<4"

    [int]$num = 0
    [int]$num2 = 0
    [int]$num8 = -1
    [int]$count = -1
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($bsString)
    [void]$sb.AppendLine('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">')
    [int]$num3 = $html.LastIndexOf('<!--EndFragment-->', [System.StringComparison]::OrdinalIgnoreCase)
    [int]$index = [Regex]::Match($html, '<\s*h\s*t\s*m\s*|', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Index
    if ($index -gt 0)
    {
        $count = $html.IndexOf('>', $index) + 1
    }
    [int]$num6 = [Regex]::Match($html, '<\s*/\s*h\s*t\s*m\s*|', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Index
    if (($html.IndexOf('<!--StartFragment-->', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -or ($num3 -ge 0))
    {
        $html
        return
    }
    [int]$startIndex = [Regex]::Match($html, '<\s*b\s*o\s*d\s*y', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Index
    if ($startIndex -gt 0)
    {
        $html.IndexOf('>', $startIndex) + 1
    }
    if (($count -lt 0) -and ($num8 -lt 0))
    {
        [void]$sb.Append('<html><body>')
        [void]$sb.Append('<!--StartFragment-->')
        $num = Get-ByteCount -sb $sb -start 0 -end -1
        [void]$sb.Append('<!--EndFragment-->')
        [void]$sb.Append('</body></html>')
    }
    else
    {
        [int]$num9 = [Regex]::Match($html, '<\s*/\s*b\s*o\s*d\s*y', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Index
        if ($count -lt 0)
        {
            [void]$sb.Append('<html>')
        }
        else
        {
            [void]$sb.Append($html, 0, $count)
        }

        if ($num8 -gt -1)
        {
            $appendThis = $count -gt -1 ? $coutn : 0
            [void]$sb.Append($html, $appendThis, ($num8 - $appendThis))
        }
        [void]$sb.Append('<!--StartFragment-->')
        $num = Get-ByteCount -sb $sb -start 0 -end -1
        if ($num8 -gt -1)
        {
            [int]$num10 = $num8
        }
        elseif ($count -gt -1)
        {
            [int]$num10 = $count
        }
        else
        {
            [int]$num10 = 0
        }

        if ($num9 -gt 0)
        {
            [int]$num11 = $num9
        }
        elseif ($num6 -gt 0)
        {
            [int]$num11 = $num6
        }
        else
        {
            [int]$num11 = $html.Length
        }

        [void]$sb.Append('<!--EndFragment-->')
        if ($num11 -lt $html.Length)
        {
            [void]$sb.Append($html, $num11, ($html.Length - $num11))
        }
        if ($num6 -le 0)
        {
            [void]$sb.Append('</html>')
        }
    }
    $sb.Replace('<<<<<<<<1', $bsString.Length.ToString("D9", [cultureinfo]::CreateSpecificCulture("en-US")), 0, $bsString.Length)
    $sb.Replace('<<<<<<<<2', ((Get-ByteCount -sb $sb -start 0 -end -1).ToString("D9", [cultureinfo]::CreateSpecificCulture("en-US"))), 0, $bsString.Length)
    $sb.Replace('<<<<<<<<3', $num.ToString("D9", [cultureinfo]::CreateSpecificCulture("en-US")), 0, $bsString.Length)
    $sb.Replace('<<<<<<<<4', $num2.ToString("D9", [cultureinfo]::CreateSpecificCulture("en-US")), 0, $bsString.Length)
    $sb.ToString()
}

Function Set-ClipboardContent([System.Collections.Generic.List[string]]$contentList, [bool]$append, [bool]$asHtml)
{
    if (($null -eq $contentList -or $contentList.Count -eq 0) -and -not $append)
    {
        if ($PSCmdlet.ShouldProcess("Clear clipboard", "Set-Clipboard"))
        {
            [System.Windows.Forms.Clipboard]::Clear()
        }
    }
    else
    {
        $builder = New-Object System.Text.StringBuilder
        if ($append)
        {
            if ([System.Windows.Forms.Clipboard]::ContainsText())
            {
                [void]$builder.AppendLine([System.Windows.Forms.Clipboard]::GetText())
            }
            else
            {
                Write-Verbose "No appendable clipboard found."
                $append = $false
            }
        }

        if ($null -ne $contentList)
        {
            [void]$builder.Append([string]::Join("`n", $contentList.ToArray(), 0, $contentList.Count))
        }

        if ($null -ne $contentList)
        {
            $str2 = $contentList[0]
            if ($str2.Length -ge 20)
            {
                $str2 = $str2.Substring(0, 20) + "..."
            }
        }

        if ($append)
        {
            $msg = "Append string '{0}' to the clipboard." -f $str2
        }
        else
        {
            $msg = "Set string '{0}' to the clipboard." -f $str2
        }

        if ($PSCmdlet.ShouldProcess($msg, "Set-Clipboard"))
        {
            [System.Windows.Forms.Clipboard]::Clear()
            if ($asHtml)
            {
                [System.Windows.Forms.Clipboard]::SetText((Get-HtmlDataString -html $builder.ToString()), [System.Windows.Forms.TextDataFormat]::Html)
            }
            else
            {
                [System.Windows.Forms.Clipboard]::SetText($builder.ToString())
            }
        }
    }
}

#endregion

#region 'THE' FUNCTION
Function Set-Clipboard()
{
    [CmdletBinding(SupportsShouldProcess=$true, DefaultParameterSetName="String")]
	[Alias("scb")]
    param
    (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = "Value")]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [string[]] $Value,

        [Parameter(Mandatory=$false)]
        [switch] $Append,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName, ParameterSetName = "Path")]
        [ValidateNotNullOrEmpty()]
        [string[]] $Path,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName, ParameterSetName = "LiteralPath")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty()]
        [string[]] $LiteralPath,

        [Parameter(Mandatory=$false)]
        [switch] $AsHtml
    )
    Begin
    {
        $contentList = New-Object -TypeName 'System.Collections.Generic.List[string]'
    }
    Process
    {
        if ($null -eq $Value -and $PSBoundParameters.ContainsKey("AsHtml"))
        {
            $PSCmdlet.ThrowTerminatingError([System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new("Html cannot be combined with the Value parameter."),
                "FailedToSetClipboard",
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                "Clipboard"
            ))
        }

        if ($PSBoundParameters.ContainsKey("Value"))
        {
            $contentList.AddRange($Value)
        }
        elseif ($PSBoundParameters.ContainsKey("Path"))
        {
            $contentList.AddRange($Path)
        }
        elseif ($PSBoundParameters.ContainsKey("LiteralPath"))
        {
            $contentList.AddRange($LiteralPath)
        }
    }
    End
    {
        if ($null -ne $LiteralPath)
        {
            Copy-FilesToClipboard -fileList $contentList -append $Append.ToBool() -isLiteralPath $true
        }
        elseif ($null -ne $Path)
        {
            Copy-FilesToClipboard -fileList $contentList -append $Append.ToBool() -isLiteralPath $false
        }
        else
        {
            Set-ClipboardContent -contentList $contentList -append $Append.ToBool() -asHtml $AsHtml.ToBool()
        }
    }
}

#endregion