Attribute VB_Name = "Module11"
Sub JobCreation()

 Dim newjobName As String
 
 'Step 1 : collect job name
 newjobName = InputBox("Enter Job Name")
 
 'Step 2 : Create job folder onto outlook folder
 'Step 2.1 check wheter jobs folder exist, if not, create it
 Dim ns As Outlook.NameSpace
 Set ns = Application.GetNamespace("MAPI")
 Dim inboxFolder As Outlook.Folder
 Set inboxFolder = ns.GetDefaultFolder(olFolderInbox)
 Dim jobsFolder As Outlook.Folder
 Dim newJobFolder As Outlook.Folder
 Dim corFolder As Outlook.Folder
 
 Set jobsFolder = AddOrUpdateFolder(inboxFolder, "Jobs")
 
 'Step 2.2 check whether new job folder exists, if not, create template
 If newjobName = "" Then Exit Sub
 Set newJobFolder = AddOrUpdateFolder(jobsFolder, newjobName)
 
 'Step 2.3 Add the templates folders here
 'Call AddOrUpdateFolder(newJobFolder, "00 not used")
 'Call AddOrUpdateFolder(newJobFolder, "01 Budget & Scope")
 Set corFolder = AddOrUpdateFolder(newJobFolder, "02 Correspondence")
 'Call AddOrUpdateFolder(newJobFolder, "03 Design Input Information")
 'Call AddOrUpdateFolder(newJobFolder, "04 Design Output")
 'Call AddOrUpdateFolder(newJobFolder, "05 Subcontractors")
 'Call AddOrUpdateFolder(newJobFolder, "06 Job Management")
 'Call AddOrUpdateFolder(newJobFolder, "07 Commissioning")
 'Call AddOrUpdateFolder(newJobFolder, "08 Financial")
 'Call AddOrUpdateFolder(newJobFolder, "09 Photographs")
 
 'Step 2.3.1 sub folders for correspondence
 Call AddOrUpdateFolder(corFolder, "Client")
 Call AddOrUpdateFolder(corFolder, "Internal")
 Call AddOrUpdateFolder(corFolder, "Suppliers-Subcon")
  
 
 'Step 3 : Create job folder and template onto disk
 Dim networkDriver As String
 'networkDriver = "C:\Jobs\" 'subject to change
 networkDriver = "R:\" 'subject to change - Auckland Intellex networkdriver..
 Call CreateFolderIfNotExists(networkDriver)
 Dim jobFolderPath As String
 jobFolderPath = networkDriver & newjobName & "\"
 Call CreateFolderIfNotExists(jobFolderPath)
 
 On Error Resume Next
 
 Dim subPath As String
 subPath = jobFolderPath & "00 not used" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "01 Budget & Scope" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "02 Correspondence" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "03 Design Input Information" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "04 Design Output" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "05 Subcontractors" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "06 Job Management" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "07 Commissioning" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "08 Financial" & "\"
  Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "09 Photographs" & "\"
  Call CreateFolderIfNotExists(subPath)
 'sub folders for correspondence
 subPath = jobFolderPath & "02 Correspondence" & "\Client\"
 Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "02 Correspondence" & "\Internal\"
 Call CreateFolderIfNotExists(subPath)
 subPath = jobFolderPath & "02 Correspondence" & "\Suppliers-Subcon\"
 Call CreateFolderIfNotExists(subPath)
 
 'Step 3, all done - force a restart of outlook
 MsgBox ("New job created, outlook needs to be restarted.")
 Application.Quit
End Sub


Function AddOrUpdateFolder(parent As Outlook.Folder, newFolderName As String) As Outlook.Folder
'check if folder exists
 On Error GoTo ErrorHandler
 Dim parentFolder As Outlook.Folder

 Set parentFolder = parent.Folders.Add(newFolderName)

 Set AddOrUpdateFolder = parentFolder
ErrorHandler:
 'folder already exists, set the current folder to this..
 Set parentFolder = parent.Folders.Item(newFolderName)
 Set AddOrUpdateFolder = parentFolder
End Function

Public Sub CreateFolderIfNotExists(folderName As String)
    'Parameter folderName must be a fully qualifed path including drive
    'All errors are assumed to be handled by the calling code
    
    Dim fs As Object 'Using late binding to avoid having to include a reference to Microsoft Scripting Runtime
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not fs.FolderExists(folderName) Then
        fs.CreateFolder (folderName)    'Comment 6
    End If
End Sub
