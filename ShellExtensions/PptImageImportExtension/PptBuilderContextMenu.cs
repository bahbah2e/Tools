namespace PptImageImportExtension
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Runtime.InteropServices;

    using ContextMenuExtension;

    [Guid("A59463CA-C047-4C74-AC1F-D9E6E2BA2DD6")]
    [ComVisible(true)]
    public class PptBuilderContextMenu : ContextMenuExt
    {
        private const string FriendlyName = "PNG To PPT";

        public PptBuilderContextMenu()
        {
            this.MenuText = "&Add to Powerpoint";
            this.Verb = "buildPPT";
            this.VerbHelpText = "Add to Powerpoint";
            this.VerbCanonicalName = "PptBuilder";
            this.SupportedFileTypes.Add(".png");
            this.SupportedFileTypes.Add(".jpg");
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
            var files = new List<FileInfo>();

            foreach (string file in this.SelectedFiles)
            {
                files.Add(new FileInfo(file));
            }

            var converter = new PptBuilder();
            converter.AddToPpt(files.ToArray());
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