namespace ContextMenuExtension
{
    using System;

    using Microsoft.Win32;

    public class ShellExtensionRegistry
    {
        private const string BadClassId = "clsid must not be empty";
        private const string BadFileType = "fileType must not be null or empty";
        private const string MenuHandlerTemplate = @"{0}\shellex\ContextMenuHandlers\{1}";
        private const string ApprovedRegistryKey = "Software\\" +
                                                   "Microsoft\\" +
                                                   "Windows\\" +
                                                   "CurrentVersion\\" +
                                                   "Shell Extensions\\" +
                                                   "Approved";

        /// <summary>
        /// Register the context menu handler.
        /// </summary>
        /// <param name="clsid">The CLSID of the component.</param>
        /// <param name="fileType">
        /// The file type that the context menu handler is associated with. For 
        /// example, '*' means all file types; '.txt' means all .txt files. The 
        /// parameter must not be NULL or an empty string. 
        /// </param>
        /// <param name="friendlyName">The friendly name of the component.</param>
        public static void RegisterShellExtContextMenuHandler(Guid clsid,
            string fileType, string friendlyName)
        {
            CheckInputs(clsid, fileType);
            AddApprovedKey(clsid, friendlyName);
            AddFileTypeRegistryKey(clsid, fileType, friendlyName);
            IssueExplorerChangeNotification();
        }

        private static void AddFileTypeRegistryKey(Guid clsid, string fileType, string friendlyName)
        {
            // Create the key HKCR\<File Type>\shellex\ContextMenuHandlers\{<CLSID>}.
            string keyName = string.Format(MenuHandlerTemplate, GetEffectiveFileType(fileType), friendlyName);
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(keyName))
            {
                // Set the default value of the key.
                if (key != null && !string.IsNullOrEmpty(friendlyName))
                {
                    Console.WriteLine(keyName + friendlyName);
                    string guid = ClassIdToString(clsid);
                    key.SetValue(null, guid);
                }
            }
        }
        private static string GetEffectiveFileType(string fileType)
        {
            // If fileType starts with '.', try to read the default value of the 
            // HKCR\<File Type> key which contains the ProgID to which the file type 
            // is linked.
            if (fileType.StartsWith("."))
            {
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(fileType))
                {
                    if (key != null)
                    {
                        // If the key exists and its default value is not empty, use 
                        // the ProgID as the file type.
                        string defaultVal = key.GetValue(null) as string;
                        if (!string.IsNullOrEmpty(defaultVal))
                        {
                            return defaultVal;
                        }
                    }
                }
            }

            return fileType;
        }
        private static void AddApprovedKey(Guid clsid, string friendlyName)
        {
            string guid = clsid.ToString("B");

            using (RegistryKey keyApproved = Registry.LocalMachine.OpenSubKey(ApprovedRegistryKey, true))
            {
                keyApproved.SetValue(guid, friendlyName);
            }
        }
        /// <summary>
        /// Unregister the context menu handler.
        /// </summary>
        /// <param name="clsid">The CLSID of the component.</param>
        /// <param name="fileType">
        /// The file type that the context menu handler is associated with. For 
        /// example, '*' means all file types; '.txt' means all .txt files. The 
        /// parameter must not be NULL or an empty string. 
        /// </param>
        public static void UnregisterShellExtContextMenuHandler(Guid clsid,
            string fileType, string friendlyName)
        {
            CheckInputs(clsid, fileType);

            RemoveGuidApprovedKey(clsid);
            RemoveContextMenuHandlerKey(fileType, friendlyName);
            IssueExplorerChangeNotification();
        }

        private static void IssueExplorerChangeNotification()
        {
            NativeMethods.SHChangeNotify(NativeMethods.SHCNE_ASSOCCHANGED, 0, IntPtr.Zero, IntPtr.Zero);
        }

        private static void RemoveContextMenuHandlerKey(string fileType, string friendlyName)
        {
            // Remove the key HKCR\<File Type>\shellex\ContextMenuHandlers\{<CLSID>}.
            string keyName = string.Format(MenuHandlerTemplate, GetEffectiveFileType(fileType), friendlyName);
            Registry.ClassesRoot.DeleteSubKeyTree(keyName, false);
        }

        private static void RemoveGuidApprovedKey(Guid clsid)
        {
            using (RegistryKey keyApproved = Registry.LocalMachine.OpenSubKey(ApprovedRegistryKey, true))
            {
                string guid = ClassIdToString(clsid);
                keyApproved.DeleteValue(guid);
            }
        }

        private static string ClassIdToString(Guid clsid)
        {
            return clsid.ToString("B");
        }

        private static void CheckInputs(Guid clsid, string fileType)
        {
            if (clsid == null)
            {
                throw new ArgumentException(BadClassId);
            }

            if (string.IsNullOrEmpty(fileType))
            {
                throw new ArgumentException(BadFileType);
            }
        }
    }
}