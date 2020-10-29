using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;

namespace EmailAutoSaver
{
    public partial class IntellexRibbon
    {
        private void IntellexRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnAddProject_Click(object sender, RibbonControlEventArgs e)
        {
            CreateInboxFolders("Current Projects");
        }

        private void CreateInboxFolders(string projFolder)
        {
            var newNameFrm = new NotificationFrm
            {
                Text = "Enter/Paste Project Name"
            };
            newNameFrm.ShowDialog();
            // Step 1 : collect job name
            // Step 2 : Create job folder onto outlook folder
            // Step 2.1 check wheter jobs folder exist, if not, create it
            Folder inbox = (Folder)Globals.ThisAddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            string newProjName = newNameFrm.txtValue.Text.Trim();
            var currentProjFolder = GlobalVars.AddOrUpdateFolder(inbox, projFolder);
            if (string.IsNullOrEmpty(newProjName)) return;
            try
            {
                var newProjFolder = GlobalVars.AddOrUpdateFolder(currentProjFolder, newProjName);
                // Step 2.3 Add the templates folders here
                var correnspondenceFolder = GlobalVars.AddOrUpdateFolder(newProjFolder, "02 Correspondence");
                GlobalVars.AddOrUpdateFolder(correnspondenceFolder, "Client");
                GlobalVars.AddOrUpdateFolder(correnspondenceFolder, "Internal");
                GlobalVars.AddOrUpdateFolder(correnspondenceFolder, "Suppliers-Subcon");
                // Step 3 : Create job folder and template onto disk
                var projectFolderPath = Path.Combine(GlobalVars.Archived_Project_DRIVER, newProjName);
                CreateFolder(projectFolderPath);
                CreateFolder(Path.Combine(projectFolderPath, "00 not used"));
                CreateFolder(Path.Combine(projectFolderPath, "01 Budget & Scope"));
                CreateFolder(Path.Combine(projectFolderPath, "02 Correspondence"));
                CreateFolder(Path.Combine(projectFolderPath, "03 Design Input Information"));
                CreateFolder(Path.Combine(projectFolderPath, "04 Design Output"));
                CreateFolder(Path.Combine(projectFolderPath, "05 Subcontractors"));
                CreateFolder(Path.Combine(projectFolderPath, "06 Job Management"));
                CreateFolder(Path.Combine(projectFolderPath, "07 Commissioning"));
                CreateFolder(Path.Combine(projectFolderPath, "08 Financial"));
                CreateFolder(Path.Combine(projectFolderPath, "09 Photographs"));
                CreateFolder(Path.Combine(projectFolderPath, "02 Correspondence", "Client"));
                CreateFolder(Path.Combine(projectFolderPath, "02 Correspondence", "Internal"));
                CreateFolder(Path.Combine(projectFolderPath, "02 Correspondence", "Suppliers-Subcon"));
                // Step 4: Trigger a re-hooking of the event handlers for new folders
                Globals.ThisAddIn.LoadEventHandlers();
            }
            catch (System.Exception er)
            {
                MessageBox.Show(er.Message.ToString());
            }
        }

        private void CreateFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path); // create if not exists..
            }
        }

        private void btnReLoad_Click(object sender, RibbonControlEventArgs e)
        {
            // Step 4: Trigger a re-hooking of the event handlers for new folders
            Globals.ThisAddIn.LoadEventHandlers();
        }

        private void btnAddArchive_Click(object sender, RibbonControlEventArgs e)
        {
            CreateInboxFolders("Archived Projects");
        }
    }
}
