Dim objShell : Set objShell = CreateObject("Shell.Application")
Dim objFolder : Set objFolder = objShell.BrowseForFolder(0, "Select the folder containing eml-files", 0)

Dim Item
Dim i : i = 0
If (NOT objFolder is nothing) Then
  Set WShell = CreateObject("WScript.Shell")
  Set objOutlook = CreateObject("Outlook.Application")
  Set Folder = objOutlook.Session.PickFolder
  If NOT Folder Is Nothing Then
    For Each Item in objFolder.Items
      If Right(Item.Name, 4) = ".eml" AND Item.IsFolder = False Then
        Set objPost = Folder.Items.Add(6)
        Set objSafePost = CreateObject("Redemption.SafePostItem")
        objSafePost.Item = objPost
        objSafePost.Import Item.Path, 1024
        objSafePost.MessageClass = "IPM.Note"
        ' remove IPM.Post icon
	Set utils = CreateObject("Redemption.MAPIUtils")
	PrIconIndex = &H10800003
        utils.HrSetOneProp objSafePost, PrIconIndex, 256, true 'Also saves the message
	i = i + 1
      End If
    Next

    MsgBox "Import completed." & vbNewLine & vbNewLine & "Imported eml-files: " & i & _
	   vbNewLine & "Imported into: " & Folder.FolderPath, 64, "Import EML"

    Set Folder = Nothing
  Else
    MsgBox "Import canceled.", 64, "Import EML"
  End If
Else
  MsgBox "Import canceled.", 64, "Import EML"
End If

Set objFolder = Nothing
Set objShell = Nothing