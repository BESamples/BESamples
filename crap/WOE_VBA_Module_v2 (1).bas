'=============================================================
' WOE Request Tracker - Enhanced VBA Module v2.0
' FEATURES:
'   - Auto-stamps Date and Completed On when Customer Name entered
'   - Auto-calculates Delete By date (+14 days) for unapproved rows
'   - Color codes rows: Green=Complete, Red=Denied/Overdue, Yellow=Waiting
'   - Bright RED highlight when Delete By date has passed
'   - Amber WARNING when Delete By is within 3 days
'   - SyncToMasterSheet utility
'   - AddNewMonthSheet utility
'=============================================================

'-------------------------------------------------------------
' PASTE THIS INTO EACH MONTHLY SHEET MODULE
' (Jan 26, Feb 26, Mar 26, etc.)
' To open: Alt+F11 > double-click the sheet name on the left
'-------------------------------------------------------------

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim watchRange As Range
    
    ' Watch Customer Name column B
    Set watchRange = Me.Range("B2:B500")
    
    Application.EnableEvents = False
    On Error GoTo SafeExit
    
    ' --- Handle Customer Name entry (Column B) ---
    If Not Intersect(Target, watchRange) Is Nothing Then
        For Each cell In Intersect(Target, watchRange)
            If Trim(cell.Value) <> "" Then
                ' Auto-stamp Date (Col A) if empty
                If Me.Cells(cell.Row, "A").Value = "" Then
                    Me.Cells(cell.Row, "A").Value = Date
                    Me.Cells(cell.Row, "A").NumberFormat = "mm/dd/yyyy"
                End If
                ' Auto-stamp Completed On (Col I) if empty
                If Me.Cells(cell.Row, "I").Value = "" Then
                    Me.Cells(cell.Row, "I").Value = Date
                    Me.Cells(cell.Row, "I").NumberFormat = "mm/dd/yyyy"
                End If
            Else
                ' Clear dates if Customer Name is cleared
                Me.Cells(cell.Row, "A").ClearContents
                Me.Cells(cell.Row, "I").ClearContents
                Me.Cells(cell.Row, "K").ClearContents
            End If
        Next cell
    End If
    
    ' --- Recalculate Delete By + colors whenever anything changes ---
    Call RefreshSheet(Me)
    
SafeExit:
    Application.EnableEvents = True
End Sub

'-------------------------------------------------------------
' PASTE THIS INTO A NEW MODULE (Insert > Module)
' Name it: modWOEHelpers
'-------------------------------------------------------------

Sub RefreshSheet(ws As Worksheet)
    ' Recalculates Delete By dates and applies all color coding for one sheet
    Dim i As Long
    Dim lastRow As Long
    Dim dateVal As Date
    Dim approvedVal As String
    Dim completeVal As String
    Dim deleteDate As Date
    Dim daysLeft As Long
    Dim needsDelete As Boolean
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim altRow As Integer
    altRow = 0
    
    For i = 2 To lastRow
        Dim custName As String
        custName = Trim(ws.Cells(i, 2).Value)
        
        ' Skip spacer rows
        If custName = "" Then
            ws.Rows(i).Interior.ColorIndex = xlNone
            GoTo NextRow
        End If
        
        altRow = altRow + 1
        approvedVal = LCase(Trim(ws.Cells(i, 6).Value))   ' Col F: WOE Request Approved
        completeVal = LCase(Trim(ws.Cells(i, 8).Value))   ' Col H: WOE Complete
        
        ' Determine if row needs a Delete By date
        needsDelete = (approvedVal = "waiting" Or approvedVal = "no" Or approvedVal = "") _
                      And completeVal <> "complete" And completeVal <> "yes" _
                      And completeVal <> "denied" And completeVal <> "cancelled"
        
        ' --- Calculate Delete By (Col K) ---
        If needsDelete Then
            If IsDate(ws.Cells(i, 1).Value) Then
                dateVal = CDate(ws.Cells(i, 1).Value)
                deleteDate = dateVal + 14
                ws.Cells(i, 11).Value = deleteDate
                ws.Cells(i, 11).NumberFormat = "mm/dd/yyyy"
                ws.Cells(i, 11).HorizontalAlignment = xlCenter
                daysLeft = deleteDate - Date
            Else
                ws.Cells(i, 11).ClearContents
                daysLeft = 999
            End If
        Else
            ws.Cells(i, 11).ClearContents
            daysLeft = 999
        End If
        
        ' --- Apply row color ---
        Dim rowColor As Long
        
        If needsDelete And ws.Cells(i, 11).Value <> "" Then
            If daysLeft < 0 Then
                ' OVERDUE - bright red
                rowColor = RGB(255, 0, 0)
                ws.Rows(i).Font.Bold = True
                ws.Rows(i).Font.Color = RGB(255, 255, 255)
            ElseIf daysLeft <= 3 Then
                ' WARNING - amber
                rowColor = RGB(255, 192, 0)
                ws.Rows(i).Font.Bold = True
                ws.Rows(i).Font.Color = RGB(0, 0, 0)
            Else
                ' Pending/Waiting - yellow
                rowColor = RGB(255, 235, 156)
                ws.Rows(i).Font.Bold = False
                ws.Rows(i).Font.Color = RGB(0, 0, 0)
            End If
        ElseIf completeVal = "complete" Or completeVal = "yes" Then
            rowColor = RGB(198, 239, 206)   ' Green
            ws.Rows(i).Font.Bold = False
            ws.Rows(i).Font.Color = RGB(0, 0, 0)
        ElseIf completeVal = "denied" Or completeVal = "cancelled" Then
            rowColor = RGB(255, 199, 206)   ' Pink/Red
            ws.Rows(i).Font.Bold = False
            ws.Rows(i).Font.Color = RGB(0, 0, 0)
        Else
            If altRow Mod 2 = 0 Then
                rowColor = RGB(214, 228, 240)  ' Light blue
            Else
                rowColor = RGB(255, 255, 255)  ' White
            End If
            ws.Rows(i).Font.Bold = False
            ws.Rows(i).Font.Color = RGB(0, 0, 0)
        End If
        
        ws.Rows(i).Interior.Color = rowColor
        
NextRow:
    Next i
End Sub

Sub RefreshAllSheets()
    ' Run RefreshSheet on every monthly tab - use this button on your dashboard
    Dim monthSheets As Variant
    monthSheets = Array("Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26")
    
    Dim shName As Variant
    For Each shName In monthSheets
        If SheetExists(CStr(shName)) Then
            Call RefreshSheet(ThisWorkbook.Sheets(CStr(shName)))
        End If
    Next shName
    
    MsgBox "All sheets refreshed! Delete By dates and colors updated.", vbInformation, "WOE Tracker"
End Sub

Sub ShowDeleteAlerts()
    ' Pops up a summary of all rows that are overdue or due within 3 days
    Dim monthSheets As Variant
    monthSheets = Array("Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26")
    
    Dim msg As String
    Dim alertCount As Integer
    alertCount = 0
    msg = "ACCOUNTS NEEDING DELETION:" & vbCrLf & String(40, "-") & vbCrLf
    
    Dim shName As Variant
    For Each shName In monthSheets
        If SheetExists(CStr(shName)) Then
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets(CStr(shName))
            Dim i As Long
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            
            For i = 2 To lastRow
                If ws.Cells(i, 11).Value <> "" And IsDate(ws.Cells(i, 11).Value) Then
                    Dim daysLeft As Long
                    daysLeft = CDate(ws.Cells(i, 11).Value) - Date
                    
                    If daysLeft <= 3 Then
                        alertCount = alertCount + 1
                        Dim status As String
                        If daysLeft < 0 Then
                            status = "OVERDUE by " & Abs(daysLeft) & " day(s)"
                        ElseIf daysLeft = 0 Then
                            status = "DELETE TODAY"
                        Else
                            status = "due in " & daysLeft & " day(s)"
                        End If
                        msg = msg & "[" & shName & "] " & ws.Cells(i, 2).Value _
                              & " | Cust#" & ws.Cells(i, 3).Value _
                              & " | " & status & vbCrLf
                    End If
                End If
            Next i
        End If
    Next shName
    
    If alertCount = 0 Then
        MsgBox "No accounts due for deletion in the next 3 days.", vbInformation, "WOE Tracker"
    Else
        MsgBox msg, vbExclamation, "WOE Tracker - " & alertCount & " Alert(s)"
    End If
End Sub

Sub SyncToMasterSheet()
    Dim masterWS As Worksheet
    Dim ws As Worksheet
    Dim monthSheets As Variant
    Dim lastMasterRow As Long
    Dim lastRow As Long
    Dim i As Long
    
    monthSheets = Array("Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26")
    
    If Not SheetExists("Master Sheet(Do Not Delete)") Then
        MsgBox "Master Sheet not found!", vbCritical
        Exit Sub
    End If
    
    Set masterWS = ThisWorkbook.Sheets("Master Sheet(Do Not Delete)")
    If masterWS.Cells(masterWS.Rows.Count, 2).End(xlUp).Row > 1 Then
        masterWS.Rows("2:" & masterWS.Rows.Count).Delete
    End If
    
    lastMasterRow = 2
    
    Dim shName As Variant
    For Each shName In monthSheets
        If SheetExists(CStr(shName)) Then
            Set ws = ThisWorkbook.Sheets(CStr(shName))
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            For i = 2 To lastRow
                If Trim(ws.Cells(i, 2).Value) <> "" Then
                    ws.Rows(i).Copy Destination:=masterWS.Rows(lastMasterRow)
                    lastMasterRow = lastMasterRow + 1
                End If
            Next i
        End If
    Next shName
    
    Call RefreshSheet(masterWS)
    MsgBox "Master Sheet synced! " & (lastMasterRow - 2) & " records copied.", vbInformation, "WOE Tracker"
End Sub

Sub AddNewMonthSheet()
    Dim newSheetName As String
    newSheetName = InputBox("Enter new sheet name (e.g., Jul 26):", "New Month Sheet")
    If newSheetName = "" Then Exit Sub
    
    If SheetExists(newSheetName) Then
        MsgBox "Sheet '" & newSheetName & "' already exists!", vbWarning
        Exit Sub
    End If
    
    ' Find last monthly sheet to use as template
    Dim lastSheet As Worksheet
    Dim monthSheets As Variant
    monthSheets = Array("Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26")
    Dim shName As Variant
    For Each shName In monthSheets
        If SheetExists(CStr(shName)) Then Set lastSheet = ThisWorkbook.Sheets(CStr(shName))
    Next shName
    
    If lastSheet Is Nothing Then MsgBox "No template sheet found!", vbCritical: Exit Sub
    
    lastSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ws.Name = newSheetName
    
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lr > 1 Then ws.Rows("2:" & lr).Delete
    
    MsgBox "New sheet '" & newSheetName & "' created and ready!", vbInformation, "WOE Tracker"
End Sub

Function SheetExists(shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Sub Workbook_Open()
    ' Runs automatically when workbook opens - checks for alerts
    Call ShowDeleteAlerts
End Sub
