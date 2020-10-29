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
            
        }

        public void LoadEventHandlers()
        {
            Folder inbox = (Folder) Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            var jobFolder = GlobalVars.AddOrUpdateFolder(inbox, "Current Projects");
            //var jobFolder = inbox.Folders["Current Projects"];
            // dynamically attach an event when user drag item into the folder to trigger file saving to local disk..
            if (jobFolder == null) return;
            foreach (Folder job in jobFolder.Folders)
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
                task.ItemAdd += new ItemsEvents_ItemAddEventHandler(AddItem);
            }
        }

        private void AddItem(object item)
        {
            // Step 1. fetch folder path etc..  
            var msg = item as MailItem;
            var fdr = msg.Parent as Folder;
            var folders = fdr.FolderPath.Split('\\').ToList();
            // gets the last two job/task name
            var jobName = folders[folders.Count -3]; // get job name
            var taskName = folders[folders.Count - 2]; // get correspondence
            var subTaskName = folders.Last(); // get sub task
            // Step 2. Save email to disk
            var dir = GlobalVars.NETWORK_DRIVER; // test c:\jobs\
            try
            {
                string jobFolderPath = Path.Combine(dir, jobName);
                CreateFolder(jobFolderPath);
                if (!Directory.Exists(jobFolderPath))
                {
                    Directory.CreateDirectory(jobFolderPath); // create if not exists..
                }
                // check if folder exist
                var destPath = Path.Combine(dir, jobName, taskName);
                CreateFolder(destPath);
                var finalPath = Path.Combine(dir, jobName, taskName, subTaskName);
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
            catch (System.Exception e)
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
            //string fileName = name + (copy == 0 ? "" : "("+ copy.ToString() + ")") + ".msg" ;
            //if (File.Exists(fileName)) {
            //    copy++;
            //    fileName = GetFileName(name, copy);
            // }
            //return fileName;
        }

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
