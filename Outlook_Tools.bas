Attribute VB_Name = "Outlook_Tools"
Sub AddPictureBorder()
    'http://stackoverflow.com/questions/12420268/how-to-put-border-round-images-in-outlook-by-default with simon edit
    Set insp = Application.ActiveInspector
    If insp.CurrentItem.Class = olMail Then
        Set mail = insp.CurrentItem
        If insp.EditorType = olEditorWord Then
            Set wordActiveDocument = mail.GetInspector.WordEditor

            For Each oIshp In wordActiveDocument.InlineShapes 'in line with text
                With oIshp.Borders
                    .OutsideLineStyle = wdLineStyleSingle
                End With
            Next oIshp

            For Each oshp In wordActiveDocument.Shapes 'floating with text wraped around
                With oshp.Line
                    .Style = msoLineSingle
                End With
            Next oshp
        End If
    End If
End Sub

Sub HelpdeskNewTicket()
	On Error GoTo errHandler
	
	Dim helpdeskaddress As String
	Dim objMail As Outlook.MailItem
	Dim fldr As Outlook.MAPIFolder
	Dim strbody As String
	Dim oldmsg As String
	Dim senderaddress As String
	Dim sendFromAccount As String
	Dim addresstype As Integer
	
	helpdeskaddress = "support@bchealth.zohodesk.eu"
	MoveToFolder = "Moved to Zoho"
	
	' Set Inbox Reference
	Set Recipient = Application.GetNamespace("MAPI").CreateRecipient("bchclinical.systemsupport")
	Recipient.Resolve
	If Recipient.Resolved Then
	    Debug.Print ("Mailbox resolved")
	End If
	Set inbox = Application.GetNamespace("MAPI").GetSharedDefaultFolder(Recipient, olFolderInbox)
	
	Set objItem = GetCurrentItem()
	
	If objItem.SenderEmailType = "EX" Then
	    senderaddress = objItem.Sender.GetExchangeUser().PrimarySmtpAddress
	Else:
	    senderaddress = objItem.SenderEmailAddress
	End If
	
	Set objMail = objItem.Forward
	
	strbody = "#original_sender {" & senderaddress & "}" & vbNewLine & vbNewLine & objItem.Body
	
	objMail.To = helpdeskaddress
	objMail.Subject = objItem.Subject
	objMail.Body = strbody
	'objMail.Display
	objMail.Send
	If objItem.Class = olMail Then
	    objItem.Move (inbox.Folders(MoveToFolder))
	End If
	
	Set objItem = Nothing
	Set objMail = Nothing
	
	errHandler:
	  MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
	   VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
    Set objApp = Application
    'On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
    Case "Explorer"
    Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
    Case "Inspector"
    Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    Case Else
    End Select
End Function

Sub DisplayUsernameOfStoreForDefaultInbox()
  ' Display name of default inbox - useful for debug
  ' Ctrl+ G to see output
  Dim NS As Outlook.NameSpace
  Dim DefaultInboxFldr As MAPIFolder

  Set NS = CreateObject("Outlook.Application").GetNamespace("MAPI")
  Set DefaultInboxFldr = NS.GetDefaultFolder(olFolderInbox)

  Debug.Print DefaultInboxFldr.Parent.Name
End Sub
