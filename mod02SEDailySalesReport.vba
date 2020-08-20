Public Const sDSemailFolder As String = "\\swek-pfsx-file1\SWEK_Finance\Logistics\Local Daily Sales (ITM)\daily email\"

Public Const sDScontrolFile As String = "dailyEmailControlFile.xlsx"
Public Const sSendingEmail As String = "logistic-nordic@amgen.com"
Public Const sTempFile As String = "temp_SE.xlsx"




Public Sub RunDailySalesSE(MItem As Outlook.MailItem)

Dim wbControleFile As Excel.Workbook
Dim oXL As Excel.Application

 Dim res As Boolean

  'MsgBox "Now running the code to do SE Sales", vbOKOnly + vbExclamation, "Hello world"

'Save the attachment
    Call SaveAttachment(MItem)

'Open Excel File
    Call OpenCommandFile(oXL, wbControleFile)

'Process the reports
   res = ProcessTheReport(oXL, wbControleFile)

'Close Excel File
   Call CloseCommandFile(res, oXL, wbControleFile)
   
'Kill excel app
Call KillExcelApp(res)


End Sub


Private Function ProcessTheReport(oXL As Excel.Application, wbControleFile As Excel.Workbook) As Boolean
    Dim countrySName As String, subFolderN As String, filterField As String, filterStr As String, attFileName As String, tempAttSave As String, macroToRun As String
    Dim res As Boolean
    Dim wb As Excel.Workbook, ws As Excel.Worksheet, dbRng As Excel.Range, outWB As Excel.Workbook, outWS1 As Excel.Worksheet, outWS2 As Excel.Worksheet, outWS3 As Excel.Worksheet
    Dim mailFilterField As String, mailFilterCriteria As String
    Dim tempArr() As String, curEmailGroup As String, lastDayStr As String, lastDayDate As Date, Subj As String
    Dim copyRng As String
    Dim sMsg As String
    Dim sStatus As String
    Dim sStatusReportRecipient As String
    Dim rnStatusRecList As Excel.Range
    Dim wsEmail As Excel.Worksheet
    sStatus = "DS Success"
    
ProcessTheReport = True
   On Error GoTo ERR:
  
  sMsg = ""
  Set wb = oXL.Workbooks.Open(wbControleFile.Path & "\" & sTempFile, ReadOnly:=True)
  Set ws = wb.Worksheets(1)
  startR = wbControleFile.Sheets(1).Cells(13 + 1, 11).Value
  startC = wbControleFile.Sheets(1).Cells(13 + 1, 12).Value
  extraRowsCount = wbControleFile.Sheets(1).Cells(13 + 1, 13).Value
  ws.UsedRange.UnMerge
  ws.Activate
  ws.Cells(startR, startC).Select
  oXL.Selection.End(xlToRight).Select
  oXL.Selection.End(xlDown).Select
  lastR = oXL.ActiveCell.Row
  lastC = oXL.ActiveCell.Column
  Set dbRng = ws.Range(ws.Cells(startR, startC), ws.Cells(lastR - extraRowsCount, lastC))
  
  Set wsEmail = wbControleFile.Worksheets(3)
  Set rnStatusRecList = wsEmail.Range("StatusRecipientList")
 
  sStatusReportRecipient = rnStatusRecList.Value
   
    With wbControleFile.Sheets(1)
        countryCount = .Range("B13").Value
        For i = 1 To countryCount
            
         
            ws.AutoFilterMode = False
                        
            countrySName = .Cells(13 + i, 3).Value
            subFolderN = .Cells(13 + i, 4).Value
            mailFilterField = .Cells(13 + i, 5).Value
            mailFilterCriteria = .Cells(13 + i, 6).Value
            attFileName = .Cells(13 + i, 7).Value
            userCheck = .Cells(13 + i, 9).Value 'if not authorized, val is -1
            tempAttSave = .Cells(13 + i, 10).Value
            dateC = .Cells(13 + i, 40).Value

            'Resort the date column if needed
            ws.Sort.SortFields.Clear
            ws.Sort.SortFields.Add Key:=ws.Range(ws.Cells(startR, dateC), ws.Cells(startR, dateC)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
            With ws.Sort
              .SetRange dbRng
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlTopToBottom
              .Apply
             End With
            
            
            lastDayStr = ws.Cells(lastR - extraRowsCount, dateC).Value
            
                      
                            
            If lastDayStr <> .Cells(13 + i, 42).Value Then    ' if you haven't already received this date
                            
                    .Cells(13 + i, 42).Value = lastDayStr
                                       
                    'Sort DB Range---------------------------------------------------------------------------------------------
                     ws.Sort.SortFields.Clear
                     For sortInd = 1 To 5
                        curSortCol = .Cells(13 + i, 14 + (sortInd - 1) * 2).Value
                        curSortDirection = IIf(.Cells(13 + i, 15 + (sortInd - 1) * 2).Value = "Smallest to Largest", xlAscending, xlDescending)
                        If Len(curSortCol) > 0 And curSortCol > 0 Then
                              ws.Sort.SortFields.Add Key:=ws.Range(ws.Cells(startR, curSortCol), ws.Cells(startR, curSortCol)), SortOn:=xlSortOnValues, Order:=curSortDirection, DataOption:=xlSortNormal 'Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
                        End If
                     Next sortInd
                     With ws.Sort
                        .SetRange dbRng
                        .Header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .Apply
                     End With
                                                                      
                     'Filter DB Area-------------------------------------------------------------------------------------------
                      ws.AutoFilterMode = False
                      dbRng.AutoFilter
                     For filterInd = 1 To 5
                       curFilterCol = .Cells(13 + i, 24 + (filterInd - 1) * 3).Value
                       curContentDepFilter = .Cells(13 + i, 25 + (filterInd - 1) * 3).Value
                       If Len(curFilterCol) > 0 And curFilterCol > 0 Then
                           If curContentDepFilter = "LastRowContent" Then
                             curValFilter = ws.Cells(lastR - extraRowsCount, curFilterCol).Value
                           ElseIf curContentDepFilter = "FirstRowContent" Then
                             curValFilter = ws.Cells(startR, curFilterCol).Value
                           ElseIf curContentDepFilter = "MinContent" Then
                             curValFilter = oXL.WorksheetFunction.Min(ws.Range(ws.Cells(startR, curFilterCol), ws.Cells(startR, curFilterCol)))
                           ElseIf curContentDepFilter = "MaxContent" Then
                             curValFilter = oXL.WorksheetFunction.Max(ws.Range(ws.Cells(startR, curFilterCol), ws.Cells(startR, curFilterCol)))
                           ElseIf Len(curContentDepFilter) = 0 Or curContentDepFilter = 0 Then
                              curValFilter = .Cells(13 + i, 26 + (filterInd - 1) * 3).Value
                           End If
                             tempArr = VBA.Split(curValFilter, ";")
                           dbRng.AutoFilter Field:=curFilterCol, Criteria1:=tempArr, Operator:=xlFilterValues
                        End If
                     Next filterInd
                                    
                          
                       templateFileName = .Cells(13 + i, 41).Value
                       outputFileName = .Cells(13 + i, 44).Value
                       Call FileCopy(wbControleFile.Path & "\" & templateFileName, wbControleFile.Path & "\" & outputFileName)
                       

                           Set outWB = openWB(wbControleFile.Path & "\" & outputFileName, oXL) ' Using this command updates links
                           Set outWS1 = outWB.Worksheets("SrcDB")
                           Set outWS2 = outWB.Worksheets("SalesDetail")
                           Set outWS3 = outWB.Worksheets("Summary")
                           dbRng.Copy
                           outWB.Activate
                           outWS1.Activate
                           outWS1.Cells(1, 1).Activate
                           oXL.ActiveSheet.Paste
                           pastedRowCount = oXL.WorksheetFunction.CountA(outWS1.Range("A:A"))
                           
                           ' Update YTD sales
                           Call updateSalesEmailLink(outWB, oXL, wbControleFile)
                           
                           'Put columns in correct order
                           orderingStr = .Cells(13 + i, 39).Value
                           tempArr = VBA.Split(orderingStr, ";")
                           For tempArrIter = LBound(tempArr) To UBound(tempArr)
        '                                            outWB.Worksheets(2).Cells(1, tempArrIter + 1).Value = tempArr(tempArrIter)
                               outWS2.Cells(1, tempArrIter + 1).Value = tempArr(tempArrIter)
                           Next tempArrIter
                           colCount = UBound(tempArr) - LBound(tempArr) + 1
                           If pastedRowCount > 1 Then 'extend formulas to be safe
                               outWS2.Range(outWS2.Cells(3, 1), outWS2.Cells(pastedRowCount + 1, colCount)).Formula = _
                                   outWS2.Range(outWS2.Cells(3, 1), outWS2.Cells(3, colCount)).Formula
                           End If
                           
                           'Formatting according to the new col order
                           For colIter = 1 To colCount
                               outWS1.Cells(2, outWS2.Cells(1, colIter).Value).Copy
                               outWS2.Activate
                               outWS2.Cells(3, colIter).PasteSpecial xlPasteFormats
                           Next colIter
                           
                           'outWS2.UsedRange.ClearFormats 'no need
                           outWS2.Range(outWS2.Cells(3, 1), outWS2.Cells(3, colCount)).Copy
                           outWS2.Range(outWS2.Cells(3, 1), outWS2.Cells(3 + pastedRowCount - 1, colCount)).PasteSpecial xlPasteFormats
                           
                           
                           outWS2.Columns.AutoFit
                           
                           'Hide
                           outWS1.Visible = xlSheetHidden
                           outWB.Worksheets("Sales YTD").Visible = xlSheetHidden
                           outWS2.Range("A1").EntireRow.Hidden = True
                           outWS3.Range("A1").EntireRow.Hidden = True
                           outWS3.Range("A1").EntireColumn.Hidden = True
                           outWS3.Activate
                           
                           ' Enter date of last work day
                           outWS3.Range("A3").Value = Date - 1
                           
                           ' Break links
                           Call BreakLinks(outWB)
                           
                           'Email
                           outWB.RefreshAll
                           outWB.Save
                           curEmailGroup = .Cells(13 + i, 45).Value
                           lastDayDate = .Cells(13 + i, 43).Value
                           Subj = .Cells(13 + i, 46).Value
                           copyRng = .Cells(13 + i, 47).Value
                           sendEmailRes = NEWsendEmailFunc(oXL, wbControleFile, wbControleFile.Path & "\" & outputFileName, curEmailGroup, outWS3.Range(copyRng), lastDayDate, Subj)
                       outWB.Close True
                       Set outWB = Nothing
                       If sendEmailRes > 0 Then
                            sMsg = sMsg & Chr(10) & "Success: " & Subj
                        Else
                            sMsg = sMsg & Chr(10) & "Email Not sent: " & Subj
                            sStatus = "DS FAILURE"
                        End If
               End If
        Next i
    End With
    
    
      Call SendStatusEmail(sStatus, sMsg, sStatusReportRecipient)
      oXL.Workbooks(wb.Name).Close (False)
      Set wb = Nothing
 
    
    Exit Function
ERR:
    ProcessTheReport = False
    sStatus = "DS FAILURE"
    If Not wb Is Nothing Then Call wb.Close(False)
    If Not outWB Is Nothing Then Call outWB.Close(False)
    Call SendStatusEmail("DS FAILURE", sMsg, sStatusReportRecipient)
    
    
End Function


Private Sub SendStatusEmail(sTatus As String, sMsg As String, sStatusReportRecipient As String)

Dim objStatuskMsg As Outlook.MailItem
   
Set objStatuskMsg = Outlook.CreateItem(olMailItem)
If sStatusReportRecipient = "" Then sStatusReportRecipient = sSendingEmail
    
    With objStatuskMsg
        .To = sStatusReportRecipient
        '.cc = CCAddress
        .SentOnBehalfOfName = sSendingEmail
        .Subject = sTatus
        .Body = sMsg
        .Display
        .Send
    End With

End Sub


Private Sub SaveAttachment(MItem As Outlook.MailItem)
    Dim oAttachment As Outlook.Attachment
    Dim sSaveFolder As String
    Dim sRootFolder As String
    Dim sSaveName As String
    Dim sExtension As String
    Dim strDate As String
    
      Call MakeCheckPath(sDSemailFolder)
      For Each oAttachment In MItem.Attachments
            If DoesFileExist(sDSemailFolder, sTempFile) Then
                 'if the file exist, delete it först.
                   Call DeleteFile(sDSemailFolder, sTempFile)
            End If
            
            oAttachment.SaveAsFile sDSemailFolder & sTempFile
      Next
      
      Set MItem = Nothing

End Sub


Private Sub OpenCommandFile(oXL As Excel.Application, wbControleFile As Excel.Workbook)
  Set oXL = CreateObject("Excel.Application")
  oXL.Visible = True
  oXL.DisplayAlerts = False
  oXL.EnableEvents = False
  Set wbControleFile = openWB(sDSemailFolder & sDScontrolFile, oXL)
  
  

End Sub


Private Sub CloseCommandFile(Result As Boolean, oXL As Excel.Application, wbControleFile As Excel.Workbook)

If Not wbControleFile Is Nothing Then
  wbControleFile.Close (Result)
End If
 
 If Not oXL Is Nothing Then
  oXL.DisplayAlerts = True
  oXL.EnableEvents = True
   oXL.Quit
   Set oXL = Nothing
   End If

End Sub


Private Sub KillExcelApp(res As Boolean)
    Dim xlapp As Excel.Application

    On Error GoTo ExitRoutine
    Set xlapp = GetObject(, "Excel.Application")
    On Error GoTo 0

    If xlapp Is Nothing Then
        'No instance was running. You can create one with
        'Set xlapp = New Excel.Application
        'but in your case it doesn't sound like you need to so:
        Exit Sub
    End If

    Dim wb As Workbook
    For Each wb In xlapp.Workbooks
        wb.Close False
    Next wb

    xlapp.Quit
    Set xlapp = Nothing
    
ExitRoutine:

End Sub






'******* BELOW HELPER FUNCTIONS

Private Sub BreakLinks(wb As Excel.Workbook)
    Dim Links As Variant
    Dim i As Integer
    Links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    
    For i = 1 To UBound(Links)
        wb.BreakLink _
        Name:=Links(i), _
        Type:=xlLinkTypeExcelLinks
    Next i
End Sub

Private Function openWB(filename As String, oXL As Excel.Application) As Excel.Workbook

    Dim wb As Excel.Workbook
    Dim AlreadyOpen As Boolean
    Dim Path As String
    Dim fname As String
    If InStr(filename, "\") > 0 Then
        fname = StrReverse(filename)
        Path = StrReverse(Right(fname, InStr(fname, "\") + 1))
        fname = StrReverse(Left(fname, InStr(fname, "\") - 1))
    Else
        fname = filename
    End If
     
    AlreadyOpen = False
    For Each wb In oXL.Workbooks
        If wb.Name = fname Then
            AlreadyOpen = True
            Exit For
        End If
    Next wb
    If AlreadyOpen Then
        Set openWB = wb
    Else
        oXL.AskToUpdateLinks = False
        oXL.DisplayAlerts = False
        If Right(fname, 4) = ".csv" Then
            oXL.ActiveWorkbook.FollowHyperlink filename
            Set openWB = oXL.Workbooks(fname)
        Else
            Set openWB = oXL.Workbooks.Open(filename) ':=filename, Format:=6, delimiter:=";")
        End If
        'oXL.DisplayAlerts = True
        oXL.AskToUpdateLinks = True
    End If

End Function


Private Function ColumnLetter(ColumnNumber As Long) As String
  If ColumnNumber > 26 Then
    ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & Chr(((ColumnNumber - 1) Mod 26) + 65)
  Else
    ColumnLetter = Chr(ColumnNumber + 64)
  End If
End Function




Private Function sortRangeOnCol(ws As Excel.Worksheet, startCol As Integer, startRow As Long, colCount As Integer, rowCount As Long, sortCol As Integer)
    Dim sortColLet As String, startColLet As String, endColLet As String
    
    sortColLet = ColumnLetter(startCol + sortCol - 1)
    startColLet = ColumnLetter(startCol + 0)
    endColLet = ColumnLetter(startCol + colCount - 1)
        
    With ws
        .Sort.SortFields.Clear
        '.Sort.SortFields.Add Key:=.Range("C5:C16"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=.Range(sortColLet & (startRow + 1) & ":" & sortColLet & (startRow + rowCount - 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End With
    
    With ws.Sort
    
        .SetRange ws.Range(startColLet & startRow & ":" & endColLet & (startRow + rowCount - 1))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
End Function


Private Sub updateSalesEmailLink(outWB As Excel.Workbook, oXL As Excel.Application, wbControleFile As Excel.Workbook)
    
    Dim ws As Excel.Worksheet
    Dim folder As String
    Dim yearmonth As String
    Dim origLinkFile As String
    Dim newLinkFile As String
    Dim first As Integer
    Dim last As Integer
    Dim ttt As String
    Dim fileRoot As String
    
    yearmonth = Year(Date - Day(Date)) * 100 + Month(Date - Day(Date)) ' Year(Date - 1) & Right("0" & Month(Date - 1), 2) ' You use date-1 since we are sending yesterday's mail. We cannot have the file refer to itself
    folder = "\\swek-pfsx-file1\SWEK_Finance\Logistics\Local Daily Sales (ITM)\daily email\"
    folder = wbControleFile.Path & "\"
    Set ws = outWB.Worksheets("Sales YTD")
    fileRoot = ws.Range("fileRoot")
    
    If ws.Range("updateSalesFromEmail") Then
    
        ' Find latest file in last month
        ttt = Dir(folder & fileRoot & yearmonth & "*")
        While Len(ttt) > 0
            newLinkFile = ttt
            ttt = Dir
        Wend
        
        ' Find current link filename
        origLinkFile = ws.Range("linkMailIn").Formula
        first = InStr(origLinkFile, "[")
        last = InStr(origLinkFile, "]")
        origLinkFile = Mid(origLinkFile, first + 1, last - first - 1)
        
        oXL.ActiveWorkbook.ChangeLink Name:=folder & origLinkFile, NewName:=folder & newLinkFile, Type:=xlExcelLinks
        ws.Range("linkMailOut") = newLinkFile
    End If

End Sub


Private Function NEWsendEmailFunc(oXL As Excel.Application, wbControleFile As Excel.Workbook, attFile As String, emailGroup As String, rng As Excel.Range, lastDay As Date, Subj As String) As Integer
   On Error GoTo ERR:
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Outlook.Recipient
    Dim objOutlookAttach As Outlook.Attachment
    
    Dim Email As String
    Dim Msg As String, URL As String
    Dim x As Double
    
    
      
    
    With wbControleFile.Sheets(3)
        RecText = ""
        emailGroupInd = oXL.WorksheetFunction.Match(emailGroup, .Range("B5:K5"), 0)
        numberOfRec = .Cells(3, 1 + emailGroupInd).Value
        For i = 1 To numberOfRec
            curRec = .Cells(5 + i, 1 + emailGroupInd).Value
            RecText = RecText & IIf(i = 1, curRec, ";" & curRec)
        Next i

        'SubjSuffix = IIf(Len(.Range("O8").Value) = 0, "", .Range("O8").Value)
        'Subj = "Daily sales for " & Year(lastDay) & "-" & Month(lastDay) & "-" & Day(lastDay) & SubjSuffix
        Msg = .Range("N11").Value
    End With
        
   
    Set objOutlookMsg = Outlook.CreateItem(olMailItem)
    
    With objOutlookMsg
        .To = RecText
        '.cc = CCAddress
        .SentOnBehalfOfName = sSendingEmail
        .Subject = Subj
        .Body = Msg
        If Len(attFile) > 0 Then .Attachments.Add attFile
        '.ReadReceiptRequested = True
        .HTMLBody = RangetoHTML(rng, oXL)
        .Display
        .Send
    End With
       
  
    NEWsendEmailFunc = 1
    Exit Function
ERR:
    NEWsendEmailFunc = 0
End Function

Function RangetoHTML(rng As Excel.Range, oXL As Excel.Application)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2010
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Excel.Workbook
 
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
 
    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = oXL.Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(2, 2).PasteSpecial Paste:=8
        .Cells(2, 2).PasteSpecial xlPasteValues, , False, False
        .Cells(2, 2).PasteSpecial xlPasteFormats, , False, False
        .Rows(1).Delete
        .Columns(1).Delete
        .Cells(2, 2).Select
        oXL.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
 
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
 
    'Close TempWB
    TempWB.Close savechanges:=False
 
    'Delete the htm file we used in this function
    Kill TempFile
 
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


 
