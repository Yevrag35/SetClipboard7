using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MG.PowerShell
{
    [Cmdlet(VerbsCommon.Get, "Clipboard", HelpUri = "https://go.microsoft.com/fwlink/?LinkId=526219")]
    [Alias("gcb")]
    [OutputType(typeof(string), typeof(FileInfo), typeof(Stream))]
    public class GetClipboard : PSCmdlet
    {
        private bool isRawSet = false;
        private bool raw;
        private bool isTextFormatTypeSet = false;
        private TextDataFormat textFormat = TextDataFormat.UnicodeText;

        [Parameter(Mandatory = false)]
        public ClipboardFormat Format { get; set; }

        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public TextDataFormat TextFormatType
        {
            get => this.textFormat;
            set
            {
                this.isTextFormatTypeSet = true;
                this.textFormat = value;
            }
        }

        [Parameter(Mandatory = false)]
        public SwitchParameter Raw
        {
            get => this.raw;
            set
            {
                this.isRawSet = true;
                this.raw = value;
            }
        }

        protected override void BeginProcessing()
        {
            if ((this.Format != ClipboardFormat.Text) && this.isTextFormatTypeSet)
            {
                base.ThrowTerminatingError(new ErrorRecord(new InvalidOperationException("Failed to get clipboard content."), "FailedToGetClipboard", ErrorCategory.InvalidOperation, "Clipboard"));
            }
            if ((this.Format != ClipboardFormat.Text) && ((this.Format != ClipboardFormat.FileDropList) && this.isRawSet))
            {
                base.ThrowTerminatingError(new ErrorRecord(new InvalidOperationException("Failed to combine raw content."), "FailedToGetClipboard", ErrorCategory.InvalidOperation, "Clipboard"));
            }
            if (this.Format == ClipboardFormat.Text)
            {
                base.WriteObject(this.GetClipboardContentAsText(this.textFormat), true);
            }
            else if (this.Format == ClipboardFormat.Image)
            {
                base.WriteObject(Clipboard.GetImage());
            }
            else if (this.Format != ClipboardFormat.FileDropList)
            {
                if (this.Format == ClipboardFormat.Audio)
                {
                    base.WriteObject(Clipboard.GetAudioStream());
                }
            }
            else if (this.raw)
            {
                base.WriteObject(Clipboard.GetFileDropList(), true);
            }
            else
            {
                base.WriteObject(this.GetClipboardContentAsFileList());
            }
        }

        private List<PSObject> GetClipboardContentAsFileList()
        {
            if (!Clipboard.ContainsFileDropList())
            {
                return null;
            }

            var list = new List<PSObject>();
            foreach (string str in Clipboard.GetFileDropList())
            {
                FileInfo fi = new FileInfo(str);
                list.Add(this.WrapOutputInPSObject(fi, str));
            }
            return list;
        }
        private List<string> GetClipboardContentAsText(TextDataFormat textFormat)
        {
            if (!Clipboard.ContainsText())
            {
                return null;
            }
            var list = new List<string>();
            string text = Clipboard.GetText(textFormat);
            if (this.raw)
            {
                list.Add(text);
            }
            else
            {
                list.AddRange(text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries));
            }
            return list;
        }
        private PSObject WrapOutputInPSObject(FileInfo fi, string path)
        {
            var obj2 = new PSObject(fi);
            if (path != null)
            {
                string fullName = Directory.GetParent(path).FullName;
                MethodInfo addOrSet = obj2.GetType().GetMethod("AddOrSetProperty", BindingFlags.NonPublic | BindingFlags.Instance,
                    null, new Type[2] { typeof(string), typeof(object) }, null);

                addOrSet.Invoke(obj2, new object[2] { "PSParentPath", fullName });
                string name = fi.Name;
                addOrSet.Invoke(obj2, new object[2] { "PSChildName", name });
            }
            return obj2;
        }
    }

    public enum ClipboardFormat
    {
        Text,
        FileDropList,
        Image,
        Audio
    }
}
