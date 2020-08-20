Option Explicit
Private objNS As Outlook.NameSpace
Private WithEvents SEobjItems As Outlook.Items
Private WithEvents DKobjItems As Outlook.Items
Private WithEvents NOobjItems As Outlook.Items

Private Const SEdir As String = "SE daily sales"
Private Const DKdir As String = "DK Movianto"
Private Const NOdir As String = "NO daily sales"



Private Sub Application_Startup()
 
Dim objWatchFolderSE As Outlook.folder
Dim objWatchFolderDK As Outlook.folder
Dim objWatchFolderNO As Outlook.folder

Set objNS = Application.GetNamespace("MAPI")

'Set the folder and items to watch:
Set objWatchFolderSE = objNS.GetDefaultFolder(olFolderInbox).Folders(SEdir)
Set objWatchFolderDK = objNS.GetDefaultFolder(olFolderInbox).Folders(DKdir)
Set objWatchFolderNO = objNS.GetDefaultFolder(olFolderInbox).Folders(NOdir)

Set SEobjItems = objWatchFolderSE.Items
Set DKobjItems = objWatchFolderDK.Items
Set NOobjItems = objWatchFolderNO.Items

Set objWatchFolderSE = Nothing
Set objWatchFolderDK = Nothing
Set objWatchFolderNO = Nothing

End Sub


Private Sub SEobjItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    Call SaveSEDailySales(Item)
   
End Sub

Private Sub DKobjItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    Call SaveMoviantoData(Iteam)
    
    End Sub
    
Private Sub NOobjItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    Call SaveNODailySales(Item)

End Sub
