Sub Archive()
    Dim myItem
    For Each myItem In Application.Explorers.Item(0).Selection
        myItem.UnRead = False
        myItem.Move (Application.Explorers.Item(1).CurrentFolder().Parent.Parent.Folders("Archive"))
    Next
End Sub
Sub TagItem()
  Set myItem = Application.Explorers.Item(0).Selection.Item(1)
  Dim form As UserForm1
  Set form = New UserForm1
  form.TextBox1.Text = myItem.Categories
  form.Show
  myItem.Categories = form.TextBox1.Text
End Sub

Sub AutomatingClassification(MyMail As MailItem)
    Dim strID As String
    Dim olNS As Outlook.NameSpace
    Dim olMail As Outlook.MailItem
    'Dim myFolder As Outlook.Folders
    'Dim myDestFolder As Outlook.Folders
    
    
    strID = MyMail.EntryID
    Set olNS = Application.GetNamespace("MAPI")
    Set olMail = olNS.GetItemFromID(strID)
    ' MsgBox olMail.Subject
    
    Set myFolder = olNS.GetDefaultFolder(olFolderInbox)
    FolderName1 = "Archive"
    
    On Error Resume Next
    Set myDestFolder = myFolder.Folders(FolderName1)
    If myDestFolder Is Nothing Then
        Set myNewFolder = myFolder.Parent.Folders.Add(FolderName1)
    End If
    olMail.Move myDestFolder
    
    Set olMail = Nothing
    Set olNS = Nothing

End Sub
