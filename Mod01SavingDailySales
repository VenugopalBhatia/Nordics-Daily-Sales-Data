Public Sub SaveNODailySales(MItem As Outlook.MailItem)
    Call SaveDailySales(MItem, "NO")
    Set MItem = Nothing
End Sub

Public Sub SaveSEDailySales(MItem As Outlook.MailItem)
    Call SaveDailySales(MItem, "SE")
    Call RunDailySalesSE(MItem)
   Set MItem = Nothing
End Sub


Public Sub SaveMoviantoData(MIteam As Outlook.MailItem)
    Call HandleMoviantoData(MIteam)
    Set MItem = Nothing
End Sub


Public Sub RunTodaySEDailySales()
    Dim ns As Outlook.NameSpace
    Dim moveToFolder As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    Dim MItem As Outlook.MailItem
    
    Set ns = Application.GetNamespace("MAPI")
    Set moveToFolder = ns.GetDefaultFolder(olFolderInbox).Folders("SE daily sales")

    If moveToFolder Is Nothing Then
       MsgBox "Folder SE daily sales not found!", vbOKOnly + vbExclamation, "Move Macro Error"
    End If
    
    Dim tnow As Double


    For Each objItem In moveToFolder.Items
       If moveToFolder.DefaultItemType = olMailItem Then
          If objItem.Class = olMail Then
            If objItem.ReceivedTime > Now - 1 Then
                Set MItem = objItem
                Set objItem = Nothing
                Set moveToFolder = Nothing
                Set ns = Nothing
                Call RunDailySalesSE(MItem)
                Set MItem = Nothing
                Exit Sub
            End If
          End If
      End If
    Next

End Sub


Public Sub TestSaveDailySales()

    Dim ns As Outlook.NameSpace
    Dim moveToFolder As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    
    Set ns = Application.GetNamespace("MAPI")
    Set moveToFolder = ns.GetDefaultFolder(olFolderInbox).Folders("SE daily sales")

    If moveToFolder Is Nothing Then
       MsgBox "Folder SE daily sales not found!", vbOKOnly + vbExclamation, "Move Macro Error"
    End If
    
    Dim tnow As Double


    For Each objItem In moveToFolder.Items
       If moveToFolder.DefaultItemType = olMailItem Then
          If objItem.Class = olMail Then
            If objItem.ReceivedTime > Now - 5 Then
             Call SaveDailySales(objItem, "SE")
             End If
          End If
      End If
    Next
    
    
    
    Set moveToFolder = ns.GetDefaultFolder(olFolderInbox).Folders("NO daily sales")

    If moveToFolder Is Nothing Then
       MsgBox "Folder NO daily sales not found!", vbOKOnly + vbExclamation, "Move Macro Error"
    End If
    
  
    For Each objItem In moveToFolder.Items
       If moveToFolder.DefaultItemType = olMailItem Then
          If objItem.Class = olMail Then
            If objItem.ReceivedTime > Now - 5 Then
             Call SaveDailySales(objItem, "NO")
             End If
          End If
      End If
    Next
    
    Set moveToFolder = ns.GetDefaultFolder(olFolderInbox).Folders("DK Movianto")

    If moveToFolder Is Nothing Then
       MsgBox "Folder DK Movianto not found", vbOKOnly + vbExclamation, "Move Macro Error"
    End If
    
    For Each objItem In moveToFolder.Items
       If moveToFolder.DefaultItemType = olMailItem Then
          If objItem.Class = olMail Then
            If objItem.ReceivedTime > Now - 20 Then
             Call SaveMoviantoData(objItem)
             End If
          End If
      End If
    Next
    

End Sub

Private Sub SaveDailySales(MItem As Outlook.MailItem, co As String, Optional strWHSName As String)
    Dim oAttachment As Outlook.Attachment
    Dim oExistingFile As Object
    Dim sSaveFolder As String
    Dim sRootFolder As String
    Dim sSaveName As String
    Dim sExtension As String
    Dim strDate As String
    sRootFolder = "\\swek-pfsx-file1\SWEK_Finance\Logistics\Local Daily Sales (ITM)\DSdata"
    Dim tReceivedTime As Date
    tReceivedTime = MItem.ReceivedTime
    
    strDate = CStr(Year(tReceivedTime)) & "-" & Right("00" & CStr(Month(tReceivedTime)), 2) & "-" & Right("00" & CStr(Day(tReceivedTime)), 2)
    
    sSaveFolder = sRootFolder & "\" & strDate & " DS data" & "\"
    
         
    If strWHSName <> "" Then
        sSaveName = strWHSName & strDate
    Else
        sSaveName = co & " DailySales" & strDate
    End If
    
    
    
    Call MakeCheckPath(sSaveFolder)
        
    
    For Each oAttachment In MItem.Attachments
        sExtension = Mid(oAttachment.DisplayName, InStr(oAttachment.DisplayName, "."), Len(oAttachment.DisplayName))
        If Not DoesFileExist(sSaveFolder, sSaveName & sExtension) Then
            oAttachment.SaveAsFile sSaveFolder & sSaveName & sExtension
        Else
          Set oExistingFile = RetrieveFile(sSaveFolder, sSaveName & sExtension)
          If oAttachment.Size > oExistingFile.Size Then
            oAttachment.SaveAsFile sSaveFolder & sSaveName & sExtension
          End If
        End If
        
    Next
    
End Sub


Private Sub HandleMoviantoData(MItem As Outlook.MailItem)
    Dim oAttachment As Outlook.Attachment
    Dim sSaveFolder As String
    Dim sRootFolder As String
    Dim sSaveName As String
    Dim sExtension As String
    Dim strDate As String
    sRootFolder = "\\swek-pfsx-file1\SWEK_Finance\Logistics\Local Daily Sales (ITM)\MoviantoData"


     Dim tReceivedTime As Date
    tReceivedTime = MItem.ReceivedTime
        strDate = CStr(Year(tReceivedTime)) & "-" & Right("00" & CStr(Month(tReceivedTime)), 2) & "-" & Right("00" & CStr(Day(tReceivedTime)), 2)
    
    sSaveFolder = sRootFolder & "\" & strDate & " Movianto data" & "\"
      Call MakeCheckPath(sSaveFolder)
      For Each oAttachment In MItem.Attachments
            If Not DoesFileExist(sSaveFolder, oAttachment.DisplayName) Then
                    oAttachment.SaveAsFile sSaveFolder & oAttachment.DisplayName
            End If
      Next
      If InStr(LCase(MItem.Subject), "saleslines") > 0 Then
            If InStr(UCase(MItem.Subject), "MTD") > 0 Then
                Call SaveDailySales(MItem, "DK", "Movianto")
            End If
      End If

End Sub
