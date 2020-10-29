using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

namespace EmailAutoSaver
{
    public partial class ThisAddIn
    {
        List<Items> _taskItems = new List<Items>();
        List<Items> _archivedItems = new List<Items>();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                LoadEventHandlers();
            }
            catch (System.Exception er)
            {
                MessageBox.Show(er.ToString());
            }
            Application.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(Email_SendingValidation);
        }

        #region Email Sender
        // append additional info onto email subjects etc when user clicks 'send'
        private void Email_SendingValidation(object Item, ref bool Cancel)
        {
            MailItem mail = Item as MailItem;
            if (mail != null)
            {
                string additionalSubject = DateTime.Now.ToString("yyyyMMdd"); // Append date time to email subject
                if (!mail.Subject.Contains(additionalSubject))
                    mail.Subject = string.Format("{0} {1}", additionalSubject, mail.Subject);
            }
        }

        #endregion

        public void LoadEventHandlers()
        {
            Folder inbox = (Folder) Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            // Get folder list for attaching the event handlers
            GetCurrentProjectItems(inbox, "Current Projects");
            GetArchiveProjectItems(inbox, "Archived Projects");
        }

        // default to 3rd level down to the nested folder structure to attach the handler when dragged into
        private void GetCurrentProjectItems(Folder inbox, string folderName)
        {
            var fdr = GlobalVars.AddOrUpdateFolder(inbox, folderName);
            if (fdr == null) return;
            foreach (Folder job in fdr.Folders)
            {
                foreach (Folder task in job.Folders)
                {
                    foreach (Folder subTask in task.Folders)
                    {
                        _taskItems.Add(subTask.Items);
                    }
                }
            }
            foreach (Items task in _taskItems)
            {
                // remove the handler just in case
                task.ItemAdd -= new ItemsEvents_ItemAddEventHandler((sender) => AddItem(sender, GlobalVars.Current_Project_DRIVER)); ;
                task.ItemAdd += new ItemsEvents_ItemAddEventHandler((sender) => AddItem(sender, GlobalVars.Current_Project_DRIVER));
            }
        }

        private void GetArchiveProjectItems(Folder inbox, string folderName)
        {
            var fdr = GlobalVars.AddOrUpdateFolder(inbox, folderName);
            if (fdr == null) return;
            foreach (Folder job in fdr.Folders)
            {
                foreach (Folder task in job.Folders)
                {
                    foreach (Folder subTask in task.Folders)
                    {
                        _archivedItems.Add(subTask.Items);
                    }
                }
            }
            foreach (Items task in _archivedItems)
            {
                // remove the handler just in case
                task.ItemAdd -= new ItemsEvents_ItemAddEventHandler((sender) => AddItem(sender, GlobalVars.Archived_Project_DRIVER)); ;
                task.ItemAdd += new ItemsEvents_ItemAddEventHandler((sender) => AddItem(sender, GlobalVars.Archived_Project_DRIVER));
            }
        }
        #region Email Saver
        private void AddItem(object item, string path)
        {
            // Step 1. fetch folder path etc..  
            var msg = item as MailItem;
            var fdr = msg.Parent as Folder;
            var folders = fdr.FolderPath.Split('\\').ToList();
            // gets the last two job/task name
            var jobName = folders[folders.Count - 3]; // get job name
            var taskName = folders[folders.Count - 2]; // get correspondence
            var subTaskName = folders.Last(); // get sub task
            // Step 2. Save email to disk
            try
            {
                string jobFolderPath = Path.Combine(path, jobName);
                CreateFolder(jobFolderPath);
                if (!Directory.Exists(jobFolderPath))
                {
                    Directory.CreateDirectory(jobFolderPath); // create if not exists..
                }
                // check if folder exist
                var destPath = Path.Combine(path, jobName, taskName);
                CreateFolder(destPath);
                var finalPath = Path.Combine(path, jobName, taskName, subTaskName);
                CreateFolder(finalPath);
                // task folder exists, save the email
                var titleEditFrm = new NotificationFrm() { Text = "Edit Email Title" };
                string msgTitle = CleanupMessageTitle(msg.Subject, msg.CreationTime);
                titleEditFrm.txtValue.Text = msgTitle; // assign msg title for edit?
                titleEditFrm.ShowDialog();
                if (string.IsNullOrEmpty(titleEditFrm.txtValue.Text)) return;
                string fileName = GetFileName(finalPath + @"\" + titleEditFrm.txtValue.Text, 0);
                if (!File.Exists(fileName)) // if file doesnt exist, create new file.
                {
                    //MessageBox.Show("File Saved as " + fileName);
                    msg.SaveAs(fileName);
                }
            }
            catch (System.Exception)
            {
                //MessageBox.Show("Path not exist or network driver not connected...");
            }
        }

        private void CreateFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path); // create if not exists..
            }
        }

        private string CleanupMessageTitle(string subject, DateTime time)
        {
            var rgx = new Regex(@"\||\/|\<|\>|""|:|\*|\\|\?");
            string currentDT = time.ToString("yyyyMMdd-hhmm");
            return currentDT + "-" + rgx.Replace(subject, "");
        }

        private string GetFileName(string name, int copy)
        {
            return name + ".msg";
        }

        #endregion

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
