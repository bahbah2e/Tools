namespace ContextMenuExtension
{
    using System;
    using System.Collections.Specialized;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Text;

    [ClassInterface(ClassInterfaceType.None)]
    public abstract class ContextMenuExt : IShellExtInit, IContextMenu
    {
        private uint IDM_DISPLAY = 0;

        private IntPtr menuBmp = IntPtr.Zero;

        private StringCollection selectedFiles = new StringCollection();
        private StringCollection supportedFiletTypes = new StringCollection();

        public ContextMenuExt()
        {
        }

        ~ContextMenuExt()
        {
            if (this.menuBmp != IntPtr.Zero)
            {
                NativeMethods.DeleteObject(this.menuBmp);
                this.menuBmp = IntPtr.Zero;
            }
        }

        protected virtual IntPtr MenuBitmap
        {
            get { return this.menuBmp; }
            set { this.menuBmp = value; }
        }

        protected abstract string MenuText { get; set; }

        protected StringCollection SelectedFiles
        {
            get { return this.selectedFiles; }
            set { this.selectedFiles = value; }
        }

        protected StringCollection SupportedFileTypes
        {
            get { return this.supportedFiletTypes; }
            set { this.supportedFiletTypes = value; }
        }

        public abstract string RegistryFriendlyName { get; set; }
        protected abstract string Verb { get; set; }
        protected abstract string VerbCanonicalName { get; set; }
        protected abstract string VerbHelpText { get; set; }

        public bool IsFolderExtension
        {
            get
            {
                if (this.SupportedFileTypes.Count > 1) { return false; }

                bool isFolderExtension = false;

                foreach (string extensionType in this.SupportedFileTypes)
                {
                    if (extensionType == "Folder")
                    {
                        isFolderExtension = true;
                    }
                }

                return isFolderExtension;
            }
        }

        [ComRegisterFunction()]
        public static void Register(Type t)
        {
            try
            {
                ContextMenuExt extension = Activator.CreateInstance(t) as ContextMenuExt;

                if (extension == null) { return; }

                if (extension.IsFolderExtension)
                {
                    ShellExtensionRegistry.RegisterShellExtContextMenuHandler(t.GUID, "Folder", extension.RegistryFriendlyName);
                }
                else
                {
                    ShellExtensionRegistry.RegisterShellExtContextMenuHandler(t.GUID, "*", extension.RegistryFriendlyName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                Console.ReadLine();
                throw;  // Re-throw the exception
            }
        }

        [ComUnregisterFunction()]
        public static void Unregister(Type t)
        {
            try
            {
                ContextMenuExt extension = Activator.CreateInstance(t) as ContextMenuExt;
                if (extension == null) { return; }

                if (extension.IsFolderExtension)
                {
                    ShellExtensionRegistry.UnregisterShellExtContextMenuHandler(t.GUID, "Folder", extension.RegistryFriendlyName);
                }
                else
                {
                    ShellExtensionRegistry.UnregisterShellExtContextMenuHandler(t.GUID, "*", extension.RegistryFriendlyName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                throw;  // Re-throw the exception
            }
        }

        /// <summary>
        /// Get information about a shortcut menu command, including the help string 
        /// and the language-independent, or canonical, name for the command.
        /// </summary>
        /// <param name="idCmd">Menu command identifier offset.</param>
        /// <param name="uFlags">
        /// Flags specifying the information to return. This parameter can have one 
        /// of the following values: GCS_HELPTEXTA, GCS_HELPTEXTW, GCS_VALIDATEA, 
        /// GCS_VALIDATEW, GCS_VERBA, GCS_VERBW.
        /// </param>
        /// <param name="pReserved">Reserved. Must be IntPtr.Zero</param>
        /// <param name="pszName">
        /// The address of the buffer to receive the null-terminated string being 
        /// retrieved.
        /// </param>
        /// <param name="cchMax">
        /// Size of the buffer, in characters, to receive the null-terminated string.
        /// </param>
        public void GetCommandString(
            UIntPtr idCmd,
            uint uFlags,
            IntPtr pReserved,
            StringBuilder pszName,
            uint cchMax)
        {
            if (idCmd.ToUInt32() == this.IDM_DISPLAY)
            {
                switch ((GCS)uFlags)
                {
                    case GCS.GCS_VERBW:
                        if (this.VerbCanonicalName.Length > cchMax - 1)
                        {
                            Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                        }
                        else
                        {
                            pszName.Clear();
                            pszName.Append(this.VerbCanonicalName);
                        }
                        break;

                    case GCS.GCS_HELPTEXTW:
                        if (this.VerbHelpText.Length > cchMax - 1)
                        {
                            Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                        }
                        else
                        {
                            pszName.Clear();
                            pszName.Append(this.VerbHelpText);
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Initialize the context menu handler.
        /// </summary>
        /// <param name="pidlFolder">
        /// A pointer to an ITEMIDLIST structure that uniquely identifies a folder.
        /// </param>
        /// <param name="pDataObj">
        /// A pointer to an IDataObject interface object that can be used to retrieve 
        /// the objects being acted upon.
        /// </param>
        /// <param name="hKeyProgID">
        /// The registry key for the file object or folder type.
        /// </param>
        public void Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID)
        {
            if (pDataObj == IntPtr.Zero)
            {
                throw new ArgumentException();
            }

            FORMATETC fe = new FORMATETC();
            fe.cfFormat = (short)CLIPFORMAT.CF_HDROP;
            fe.ptd = IntPtr.Zero;
            fe.dwAspect = DVASPECT.DVASPECT_CONTENT;
            fe.lindex = -1;
            fe.tymed = TYMED.TYMED_HGLOBAL;
            STGMEDIUM stm = new STGMEDIUM();

            // The pDataObj pointer contains the objects being acted upon. In this 
            // example, we get an HDROP handle for enumerating the selected files 
            // and folders.
            IDataObject dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
            dataObject.GetData(ref fe, out stm);

            try
            {
                // Get an HDROP handle.
                IntPtr hDrop = stm.unionmember;
                if (hDrop == IntPtr.Zero)
                {
                    throw new ArgumentException();
                }

                // Determine how many files are involved in this operation.
                uint nFiles = NativeMethods.DragQueryFile(hDrop, UInt32.MaxValue, null, 0);

                // This code sample displays the custom context menu item when only 
                // one file is selected. 
                if (nFiles == 1 && this.IsFolderExtension)
                {
                    // Get the path of the file.
                    StringBuilder fileName = new StringBuilder(260);
                    if (0 == NativeMethods.DragQueryFile(hDrop, 0, fileName,
                        fileName.Capacity))
                    {
                        Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                    }
                    DirectoryInfo di = new DirectoryInfo(fileName.ToString());
                    
                    if (di.Exists)
                    {
                        this.selectedFiles.Add(fileName.ToString());
                        return;
                    }
                    else
                    {
                        Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                    }

                }
                //else
                //{
                //    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                //}

                // [-or-]

                // Enumerate the selected files and folders.
                if (nFiles > 0)
                {
                    StringBuilder fileName = new StringBuilder(260);
                    for (uint i = 0; i < nFiles; i++)
                    {
                        // Get the next file name.
                        if (0 != NativeMethods.DragQueryFile(hDrop, i, fileName, fileName.Capacity))
                        {
                            string fileNameText = fileName.ToString();
                            string fileExtension = this.GetFileExtension(fileNameText);

                            if (this.supportedFiletTypes.Contains(fileExtension))
                            {
                                // Add the file name to the list.
                                this.selectedFiles.Add(fileNameText);
                            }
                        }
                    }

                    // If we did not find any files we can work with, throw 
                    // exception.
                    if (this.selectedFiles.Count == 0)
                    {
                        Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                    }
                }
                else
                {
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }
            finally
            {
                NativeMethods.ReleaseStgMedium(ref stm);
            }
        }

        private string GetFileExtension(string fileName)
        {
            int lastDot = fileName.LastIndexOf('.');
            return fileName.Substring(lastDot);
        }
        /// <summary>
        /// Carry out the command associated with a shortcut menu item.
        /// </summary>
        /// <param name="pici">
        /// A pointer to a CMINVOKECOMMANDINFO or CMINVOKECOMMANDINFOEX structure 
        /// containing information about the command. 
        /// </param>
        public void InvokeCommand(IntPtr pici)
        {
            bool isUnicode = false;

            // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
            // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
            // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
            // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
            // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
            // and has additional members that allow Unicode strings to be passed.
            CMINVOKECOMMANDINFO ici = (CMINVOKECOMMANDINFO)Marshal.PtrToStructure(
                pici, typeof(CMINVOKECOMMANDINFO));
            CMINVOKECOMMANDINFOEX iciex = new CMINVOKECOMMANDINFOEX();
            if (ici.cbSize == Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX)))
            {
                if ((ici.fMask & CMIC.CMIC_MASK_UNICODE) != 0)
                {
                    isUnicode = true;
                    iciex = (CMINVOKECOMMANDINFOEX)Marshal.PtrToStructure(pici,
                        typeof(CMINVOKECOMMANDINFOEX));
                }
            }

            // Determines whether the command is identified by its offset or verb.
            // There are two ways to identify commands:
            // 
            //   1) The command's verb string 
            //   2) The command's identifier offset
            // 
            // If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
            // lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
            // holds a verb string. If the high-order word is zero, the command 
            // offset is in the low-order word of lpcmi->lpVerb.

            // For the ANSI case, if the high-order word is not zero, the command's 
            // verb string is in lpcmi->lpVerb. 
            if (!isUnicode && NativeMethods.HighWord(ici.verb.ToInt32()) != 0)
            {
                // Is the verb supported by this context menu extension?
                if (Marshal.PtrToStringAnsi(ici.verb) == this.Verb)
                {
                    this.OnVerb(ici.hwnd);
                }
                else
                {
                    // If the verb is not recognized by the context menu handler, it 
                    // must return E_FAIL to allow it to be passed on to the other 
                    // context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }

            // For the Unicode case, if the high-order word is not zero, the 
            // command's verb string is in lpcmi->lpVerbW. 
            else if (isUnicode && NativeMethods.HighWord(iciex.verbW.ToInt32()) != 0)
            {
                // Is the verb supported by this context menu extension?
                if (Marshal.PtrToStringUni(iciex.verbW) == this.Verb)
                {
                    this.OnVerb(ici.hwnd);
                }
                else
                {
                    // If the verb is not recognized by the context menu handler, it 
                    // must return E_FAIL to allow it to be passed on to the other 
                    // context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }

            // If the command cannot be identified through the verb string, then 
            // check the identifier offset.
            else
            {
                // Is the command identifier offset supported by this context menu 
                // extension?
                if (NativeMethods.LowWord(ici.verb.ToInt32()) == this.IDM_DISPLAY)
                {
                    this.OnVerb(ici.hwnd);
                }
                else
                {
                    // If the verb is not recognized by the context menu handler, it 
                    // must return E_FAIL to allow it to be passed on to the other 
                    // context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }
        }

        /// <summary>
        /// Add commands to a shortcut menu.
        /// </summary>
        /// <param name="hMenu">A handle to the shortcut menu.</param>
        /// <param name="iMenu">
        /// The zero-based position at which to insert the first new menu item.
        /// </param>
        /// <param name="idCmdFirst">
        /// The minimum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="idCmdLast">
        /// The maximum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="uFlags">
        /// Optional flags that specify how the shortcut menu can be changed.
        /// </param>
        /// <returns>
        /// If successful, returns an HRESULT value that has its severity value set 
        /// to SEVERITY_SUCCESS and its code value set to the offset of the largest 
        /// command identifier that was assigned, plus one.
        /// </returns>
        public int QueryContextMenu(
            IntPtr hMenu,
            uint iMenu,
            uint idCmdFirst,
            uint idCmdLast,
            uint uFlags)
        {
            // If uFlags include CMF_DEFAULTONLY then we should not do anything.
            if (((uint)CMF.CMF_DEFAULTONLY & uFlags) != 0)
            {
                return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 0);
            }

            // Use either InsertMenu or InsertMenuItem to add menu items.
            MENUITEMINFO mii = new MENUITEMINFO();
            mii.cbSize = (uint)Marshal.SizeOf(mii);
            mii.fMask = MIIM.MIIM_BITMAP | MIIM.MIIM_STRING | MIIM.MIIM_FTYPE |
                        MIIM.MIIM_ID | MIIM.MIIM_STATE;
            mii.wID = idCmdFirst + this.IDM_DISPLAY;
            mii.fType = MFT.MFT_STRING;
            mii.dwTypeData = this.MenuText;
            mii.fState = MFS.MFS_ENABLED;
            mii.hbmpItem = this.menuBmp;
            if (!NativeMethods.InsertMenuItem(hMenu, iMenu, true, ref mii))
            {
                int errorCode = Marshal.GetHRForLastWin32Error();
                return errorCode;
            }

            // Add a separator.
            MENUITEMINFO sep = new MENUITEMINFO();
            sep.cbSize = (uint)Marshal.SizeOf(sep);
            sep.fMask = MIIM.MIIM_TYPE;
            sep.fType = MFT.MFT_SEPARATOR;
            if (!NativeMethods.InsertMenuItem(hMenu, iMenu + 1, true, ref sep))
            {
                int errorCode = Marshal.GetHRForLastWin32Error();
                return errorCode;
            }

            // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
            // Set the code value to the offset of the largest command identifier 
            // that was assigned, plus one (1).
            return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0,
                this.IDM_DISPLAY + 1);
        }

        protected abstract void OnVerb(IntPtr hWnd);
    }
}
