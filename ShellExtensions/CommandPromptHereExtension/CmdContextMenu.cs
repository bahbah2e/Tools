namespace CommandPromptHereExtension
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Runtime.InteropServices;

    using ContextMenuExtension;

    [Guid("C9DE41A6-53A2-4436-B374-769E82F0C5CA")]
    [ComVisible(true)]
    public class PstContextMenu : ContextMenuExt
    {
        private const string FriendlyName = "Command Prompt Here Extension";
        public PstContextMenu()
        {
            this.MenuText = "&Command Prompt Here...";
            this.Verb = "promptHere";
            this.VerbHelpText = "Command Prompt Here";
            this.VerbCanonicalName = "PromptHere";
            this.SupportedFileTypes.Add("Folder");
        }

        public override string RegistryFriendlyName
        {
            get { return FriendlyName; }
            set { }
        }

        protected override string MenuText { get; set; }
        protected override string Verb { get; set; }
        protected override string VerbCanonicalName { get; set; }
        protected override string VerbHelpText { get; set; }

        protected override void OnVerb(IntPtr hWnd)
        {
            List<FileInfo> files = new List<FileInfo>();

            foreach (string file in this.SelectedFiles)
            {
                Process.Start("cmd", string.Format("/k cd {0}", file));
            }
        }

        private IntPtr BuildBitmap()
        {
            // Load the bitmap for the menu item.
            //Bitmap bmp = Resources.OK;
            //bmp.MakeTransparent(bmp.GetPixel(0, 0));
            //this.MenuBitmap = bmp.GetHbitmap();
            return IntPtr.Zero;
        }
    }
}