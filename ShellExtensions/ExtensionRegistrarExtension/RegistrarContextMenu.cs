namespace ExtensionRegistrarExtension
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    using ContextMenuExtension;

    [Guid("1D58B898-D0C7-4B08-BE1C-903A663A3C4E")]
    [ComVisible(true)]
    public class RegistrarContextMenu : ContextMenuExt
    {
        private const string FriendlyName = "Shell Extension Registrar";
        public RegistrarContextMenu()
        {
            this.MenuText = "&Register";
            this.Verb = "registerDll";
            this.VerbHelpText = "Register shell extension";
            this.VerbCanonicalName = "DllRegister";
            this.SupportedFileTypes.Add(".dll");
        }

        public override string RegistryFriendlyName
        {
            get { return FriendlyName; }
            set {}
        }
        
        protected override string MenuText { get; set; }

        protected override string Verb { get; set; }

        protected override string VerbCanonicalName { get; set; }

        protected override string VerbHelpText { get; set; }

        protected override void OnVerb(IntPtr hWnd)
        {
            var files = new List<FileInfo>();

            foreach (string fileName in this.SelectedFiles)
            {
                try
                {
                    Process.Start("ExtensionRegistrar", string.Format("\"{0}\" /register", fileName));
                    MessageBox.Show("Registration complete.");       
                }
                catch (Exception e)
                {
                    MessageBox.Show("Be sure ExtensionRegistrar.exe is set in the system path.");
                }
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