using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MG.PowerShell
{
    [Cmdlet(VerbsCommon.Set, "Clipboard", ConfirmImpact = ConfirmImpact.Low, SupportsShouldProcess = true, DefaultParameterSetName = "String",
        HelpUri = "https://go.microsoft.com/fwlink/?LinkId=526220")]
    [OutputType(typeof(void))]
    [CmdletBinding(PositionalBinding = false)]
    public class SetClipboard : PSCmdlet
    {
        #region FIELDS/CONSTANTS
        private List<string> list;
        private bool asHtml;
        private bool isHtmlSet;

        #endregion

        #region PARAMETERS
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "Value")]
        [AllowNull]
        [AllowEmptyCollection]
        [AllowEmptyString]
        public string[] Value { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter Append { get; set; }

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "Path")]
        [ValidateNotNullOrEmpty]
        public string[] Path { get; set; }

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty]
        public string[] LiteralPath { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter AsHtml
        {
            get => asHtml;
            set
            {
                isHtmlSet = true;
                asHtml = value;
            }
        }

        #endregion

        #region CMDLET PROCESSING
        protected override void BeginProcessing()
        {
            this.list = new List<string>();
        }

        protected override void ProcessRecord()
        {
            if (this.Value == null && this.isHtmlSet)
            {
                base.ThrowTerminatingError(
                    new ErrorRecord(
                        new InvalidOperationException("Cannot specify 'AsHtml' when no value is set."), 
                        "FailedToSetClipboard", 
                        ErrorCategory.InvalidOperation, 
                        "Clipboard"));
            }

            if (this.Value != null)
            {
                this.list.AddRange(this.Value);
            }
            else if (this.Path != null)
            {
                this.list.AddRange(this.Path);
            }
            else if (this.LiteralPath != null)
            {
                this.list.AddRange(this.LiteralPath);
            }
        }

        protected override void EndProcessing()
        {
            if (this.MyInvocation.BoundParameters.ContainsKey("LiteralPath"))
            {
                this.CopyFilesToClipboard(this.list, this.Append.ToBool(), true);
            }
            else if (this.MyInvocation.BoundParameters.ContainsKey("Path"))
            {
                this.CopyFilesToClipboard(this.list, this.Append.ToBool(), false);
            }
            else
            {
                this.SetClipboardContent(this.list, this.Append.ToBool(), this.asHtml);
            }
        }

        #endregion

        #region BACKEND METHODS
        private void CopyFilesToClipboard(List<string> fileList, bool append, bool isLiteralPath)
        {
            int count = 0;
            var source = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (append)
            {
                if (!Clipboard.ContainsFileDropList())
                {
                    base.WriteVerbose("No appendable clipboard content was found.");
                    append = false;
                }
                else
                {
                    StringCollection fileDropList = Clipboard.GetFileDropList();
                    source = new HashSet<string>(fileDropList.Cast<string>().ToList(), StringComparer.OrdinalIgnoreCase);
                    count = fileDropList.Count;
                }
            }

            for (int i = 0; i < fileList.Count; i++)
            {
                var resolvedProviderPathFromPSPath = new Collection<string>();
                try
                {
                    if (isLiteralPath)
                    {
                        resolvedProviderPathFromPSPath.Add(base.SessionState.Path.GetUnresolvedProviderPathFromPSPath(fileList[i]));
                    }
                    else
                    {
                        resolvedProviderPathFromPSPath = base.SessionState.Path.GetResolvedProviderPathFromPSPath(fileList[i], out ProviderInfo provider);
                    }
                }
                catch (ItemNotFoundException exception)
                {
                    base.WriteError(new ErrorRecord(exception, "FailedToSetClipboard", ErrorCategory.InvalidOperation, "Clipboard"));
                }

                foreach (string str2 in resolvedProviderPathFromPSPath)
                {
                    if (!source.Contains(str2))
                    {
                        source.Add(str2);
                    }
                }
            }

            if (source.Count > 0)
            {
                string msg = null;
                if ((source.Count - count) == 1)
                {
                    if (append)
                    {
                        msg = string.Format("Append single file '{0}' to the clipboard.", source.ElementAt<string>(source.Count - 1));
                    }
                    else
                    {
                        msg = string.Format("Set single file '{0}' to the clipboard.", source.ElementAt<string>(0));
                    }
                }
                else if (append)
                {
                    msg = string.Format("Append {0} files to the clipboard.", source.Count - count);
                }
                else
                {
                    msg = string.Format("Set {0} files to the clipboard.", source.Count);
                }

                if (base.ShouldProcess(msg, "Set-Clipboard"))
                {
                    Clipboard.Clear();
                    var filePaths = new StringCollection();
                    filePaths.AddRange(source.ToArray());
                    Clipboard.SetFileDropList(filePaths);
                }
            }
        }
        private int GetByteCount(StringBuilder sb, int start = 0, int end = -1)
        {
            char[] chars = new char[1];
            int num = 0;
            end = end > -1 ? end : sb.Length;
            for (int i = start; i < end; i++)
            {
                chars[0] = sb[i];
                num += Encoding.UTF8.GetByteCount(chars);
            }
            return num;
        }
        private void SetClipboardContent(List<string> contentList, bool append, bool asHtml)
        {
            string str = null;
            if ((contentList == null || contentList.Count <= 0) && !append)
            {
                str = "Clipboard cleared";
                if (base.ShouldProcess(str, "Set-Clipboard"))
                {
                    Clipboard.Clear();
                }
            }
            else
            {
                var builder = new StringBuilder();
                if (append)
                {
                    if (Clipboard.ContainsText())
                    {
                        builder.AppendLine(Clipboard.GetText());
                    }
                    else
                    {
                        base.WriteVerbose("No appendable clipboard content was found.");
                        append = false;
                    }
                }

                if (contentList != null)
                {
                    builder.Append(string.Join(Environment.NewLine, contentList.ToArray(), 0, contentList.Count));
                }
                string str2 = null;
                if (contentList != null)
                {
                    str2 = contentList[0];
                    if (str2.Length >= 20)
                    {
                        str2 = str2.Substring(0, 20) + " ...";
                    }
                }
                if (append)
                {
                    str = string.Format("Appending the following content: {0}", str2);
                }
                else
                {
                    str = string.Format("Setting the following content: {0}", str2);
                }

                if (base.ShouldProcess(str, "Set-Clipboard"))
                {
                    Clipboard.Clear();
                    if (asHtml)
                    {
                        Clipboard.SetText(this.GetHtmlDataString(builder.ToString()), TextDataFormat.Html);
                    }
                    else
                    {
                        Clipboard.SetText(builder.ToString());
                    }
                }
            }
        }
        private const string bsString = @"Version:0.9\r\nStartHTML:<<<<<<<<1\r\nEndHTML:<<<<<<<<2\r\nStartFragment:<<<<<<<<3\r\nEndFragment:<<<<<<<<4\r\nStartSelection:<<<<<<<<3\r\nEndSelection:<<<<<<<<4";
        private string GetHtmlDataString(string html)
        {
            int num = 0;
            int num2 = 0;
            int count = -1;

            var sb = new StringBuilder();
            sb.AppendLine(bsString);
            sb.AppendLine("<!DOCTYPE HTML PUBLIC \" -//W3C//DTD HTML 4.0 Transitional//EN\">");
            int num3 = html.LastIndexOf("<!--EndFragment-->", StringComparison.OrdinalIgnoreCase);
            int index = Regex.Match(html, @"<\s*h\s*t\s*m\s*|", RegexOptions.IgnoreCase).Index;

            if (index > 0)
            {
                count = html.IndexOf(">", index) + 1;
            }
            int num6 = Regex.Match(html, @"<\s*/\s*h\s*t\s*m\s*|", RegexOptions.IgnoreCase).Index;
            if (html.IndexOf("<!--StartFragment-->", StringComparison.OrdinalIgnoreCase) >= 0 || num3 > 0)
            {
                return html;
            }
            int startIndex = Regex.Match(html, @"<\s*b\s*o\s*d\s*y", RegexOptions.IgnoreCase).Index;
            int num8 = startIndex > 0 ? html.IndexOf(">", startIndex) + 1 : -1;
            if ((count < 0) && (num8 < 0))
            {
                sb.Append("<html><body>");
                sb.Append("<!--StartFragment-->");
                num = this.GetByteCount(sb);
                sb.Append(html);
                num2 = this.GetByteCount(sb);
                sb.Append("<!--EndFragment-->");
                sb.Append("</body></html>");
            }
            else
            {
                int num9 = Regex.Match(html, @"<\s*/\s*b\s*o\s*d\s*y", RegexOptions.IgnoreCase).Index;
                if (count < 0)
                {
                    sb.Append("<html>");
                }
                else
                {
                    sb.Append(html, 0, count);
                }

                if (num8 > -1)
                {
                    sb.Append(html, count > -1 ? count : 0, num8 - ((count > -1) ? count : 0));
                }
                sb.Append("<!--StartFragment-->");
                num = this.GetByteCount(sb);
                int num10 = num8 > -1 ? num8 : (count > -1 ? count : 0);
                int num11 = num9 > 0 ? num9 : (num6 > 0 ? num6 : html.Length);
                sb.Append(html, num10, num11 - num10);
                num2 = this.GetByteCount(sb);
                sb.Append("<!--EndFragment-->");
                if (num11 < html.Length)
                {
                    sb.Append(html, num11, html.Length - num11);
                }

                if (num6 <= 0)
                {
                    sb.Append("</html>");
                }
            }
            sb.Replace("<<<<<<<<1", bsString.Length.ToString("D9", CultureInfo.CreateSpecificCulture("en-US")), 0, bsString.Length);
            sb.Replace("<<<<<<<<2", this.GetByteCount(sb).ToString("D9", CultureInfo.CreateSpecificCulture("en-US")), 0, bsString.Length);
            sb.Replace("<<<<<<<<3", num.ToString("D9", CultureInfo.CreateSpecificCulture("en-US")), 0, bsString.Length);
            sb.Replace("<<<<<<<<4", num2.ToString("D9", CultureInfo.CreateSpecificCulture("en-US")), 0, bsString.Length);
            return sb.ToString();
        }

        #endregion
    }
}