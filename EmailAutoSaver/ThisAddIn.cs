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
        private List<Items> _taskItems;
        private List<Items> _archivedItems;
        //Inspectors inspectors;
        private Outlook.Application _application = null;
        public Folder inbox; // move to global scope

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                inbox = (Folder)Application.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                EmailList = new Dictionary<string, string>();
                LoadEventHandlers();
            }
            catch (System.Exception er)
            {
                MessageBox.Show(er.ToString());
            }
            // adding event handlers
            _application = Globals.ThisAddIn.Application;
            _application.ItemLoad += new ApplicationEvents_11_ItemLoadEventHandler(LoadItems);
           
            //inspectors = Application.Inspectors;
            //inspectors.NewInspector += new InspectorsEvents_NewInspectorEventHandler(Email_SendingValidation);
        }

        // append additional info onto email subjects etc when user create/reply/forward/replyall
        #region Email Sender
        // adding event hookers to items gets loaded, i.e. emails etc ==
        private void LoadItems(object item)
        {
            ItemEvents_10_Event ie = item as ItemEvents_10_Event;
            ie.ReplyAll += new ItemEvents_10_ReplyAllEventHandler(SubjectValidation);
            ie.Reply += new ItemEvents_10_ReplyEventHandler(SubjectValidation);
            ie.Forward += new ItemEvents_10_ForwardEventHandler(SubjectValidation);
        }

        //private void Email_SendingValidation(Inspector inspector)
        //{
        //    MailItem mail = inspector.CurrentItem as MailItem;
        //    AddOrUpdateSubject(mail);
        //}

        private void SubjectValidation(object item, ref bool Cancel)
        {
            MailItem mt = item as MailItem;
            AddOrUpdateSubject(mt);
        }

        private void AddOrUpdateSubject(MailItem mail)
        {
            if (mail != null)
            {
                string additionalSubject = DateTime.Now.ToString("yyyyMMdd"); // Append date time to email subject
                if (mail.Subject != null)
                {
                    if (!mail.Subject.StartsWith(additionalSubject)) // change only when the subject doesnt start with yyyyMMdd
                        mail.Subject = string.Format("{0} {1}", additionalSubject, mail.Subject);
                }
                else
                {
                    mail.Subject = additionalSubject;
                }      
            }
        }
        #endregion

        public void LoadEventHandlers()
        {
            // Get folder list for attaching the event handlers
            _taskItems = new List<Items>();
            _archivedItems = new List<Items>();
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
                var handler = new ItemsEvents_ItemAddEventHandler((sender) => AddItem(sender, GlobalVars.Current_Project_DRIVER));
                task.ItemAdd -= handler;
                task.ItemAdd += handler;
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
        private Dictionary<string, string> EmailList;
        private void AddItem(object item, string path)
        {
            // Step 1. fetch folder path etc..  
           
            var msg = item as MailItem;
            var fdr = msg.Parent as Folder;
            var folders = fdr.FolderPath.Split('\\').ToList();
            // in place to prevent a multi-firing issue when hookers are refreshed and rebinded with new handlers.
            try
            {
                EmailList.Add(msg.EntryID, "11");
            }
            catch (System.Exception)
            {
                // same email catched, exit process
                return;
            }           
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
