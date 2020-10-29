using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailAutoSaver
{
    public static class GlobalVars
    {
        public static string NETWORK_DRIVER = @"R:\";

        public static Folder AddOrUpdateFolder(Folder parent, string newFolderName)
        {
            Folder f = null;
            if (parent.Folders.Count == 0)
            {
                return (Folder)parent.Folders.Add(newFolderName);
            }
            foreach (Folder fdr in parent.Folders)
            {
                if (fdr.Name == newFolderName)
                {
                    f = fdr;
                }
            }
            if (f == null)
            {
                f = (Folder)parent.Folders.Add(newFolderName);
            }
            return f;
        }
    }
}
