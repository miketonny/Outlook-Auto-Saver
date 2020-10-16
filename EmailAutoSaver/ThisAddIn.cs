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
            Outlook.MAPIFolder inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            var jobFolder = inbox.Folders["Jobs"];
            // dynamically attach an event when user drag item into the folder to trigger file saving to local disk..
            foreach (Outlook.Folder job in jobFolder.Folders)
            {
                foreach (Outlook.Folder task in job.Folders)
                {
                    foreach (Outlook.Folder subTask in task.Folders)
                    {
                        _taskItems.Add(subTask.Items);
                    }
                }             
            }
            foreach (Items task in _taskItems)
            {
                task.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(AddItem);
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
            var dir = @"R:\"; // test c:\jobs\
            try
            {
                string jobFolderPath = Path.Combine(dir, jobName);
                //MessageBox.Show(jobFolderPath);
                if (!Directory.Exists(jobFolderPath))
                {
                    Directory.CreateDirectory(jobFolderPath); // create if not exists..
                }
                // check if folder exist
                var destPath = Path.Combine(dir, jobName, taskName);
                //MessageBox.Show(destPath);
                if (!Directory.Exists(destPath))
                {
                    Directory.CreateDirectory(destPath); // create if not exists..
                }
                var finalPath = Path.Combine(dir, jobName, taskName, subTaskName);
                // MessageBox.Show(finalPath);
                if (!Directory.Exists(finalPath))
                {
                    Directory.CreateDirectory(finalPath); // create if not exists..
                }
                // task folder exists, save the email
                string msgTitle = CleanupMessageTitle(msg.Subject, msg.CreationTime);
                var fileName = GetFileName(finalPath + @"\" + msgTitle, 0);
                msg.SaveAs(fileName);
            }
            catch (System.Exception e)
            {
                //MessageBox.Show("Path not exist or network driver not connected...");
            }
        }

        private string CleanupMessageTitle(string subject, DateTime time)
        {
            var rgx = new Regex(@"\||\/|\<|\>|""|:|\*|\\|\?");
            string currentDT = time.ToString("yyyyMMdd-hhmmss");
            return currentDT + "-" + rgx.Replace(subject, "");
        }

        private string GetFileName(string name, int copy)
        {
            string fileName = name + (copy == 0 ? "" : "("+ copy.ToString() + ")") + ".msg" ;
            if (File.Exists(fileName)) {
                copy++;
                fileName = GetFileName(name, copy);
             }
            return fileName;
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
