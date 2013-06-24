namespace SampleExtension
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Runtime.InteropServices;

    using ContextMenuExtension;

    [Guid("7D546B1B-D333-421F-B8EB-732955A06B55"), ComVisible(true)]
    public class PstContextMenu : ContextMenuExt
    {
        public PstContextMenu()
        {
            this.MenuText = "&Destroy";
            this.Verb = "destroyPst";
            this.VerbHelpText = "Destroy PST file";
            this.VerbCanonicalName = "PstDestroyer";
            this.SupportedFileTypes.Add(".pst");
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
                File.Delete(file);
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
