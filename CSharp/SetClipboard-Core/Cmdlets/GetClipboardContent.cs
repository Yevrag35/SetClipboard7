using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace MG.PowerShell.Clipboard.Cmdlets
{
    [Cmdlet(VerbsCommon.Get, "ClipboardContent")]
    public class GetClipboardContent : PSCmdlet
    {
        [Parameter]
        [Alias("Type")]
        public ClipboardContentType ContentType { get; set; } = ClipboardContentType.File;

        protected override void ProcessRecord()
        {

        }

        private void GetFileContent()
        {

        }
    }

    public enum ClipboardContentType
    {
        Audio,
        File,
        Image
    }
}
