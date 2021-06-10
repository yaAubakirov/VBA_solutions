Attribute VB_Name = "submissionsFill"
Sub submissionsFill()

Dim lRow, lCol, lRowTemp, lColTemp, term As Long
Dim cellChecked, cellCompared As Range
Dim mainWb, wb As Workbook
Dim mainWs, ws As Worksheet
Dim mDPbase, mDPcomp As String
Dim sdbM, sdbDP, sdbRev, sdbStatus As Variant

Set mainWb = ThisWorkbook
Set mainWs = mainWb.Worksheets("data")

'to turn off a screen updating and all allerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
'find last row and column. considered 10001 because formated table's cell is understood as filled cell
    lRow = mainWs.Cells(10001, 1).End(xlUp).Row
    lCol = mainWs.Cells(1, Columns.Count).End(xlToLeft).Column

'to unprotect the sheet
    mainWs.Unprotect
    
'to erase previouse data and clear formating
    mainWs.Range(mainWs.Cells(2, 2), mainWs.Cells(lRow, lCol)).ClearFormats
    mainWs.Range(mainWs.Cells(2, 2), mainWs.Cells(lRow, lCol)).Clear
    
'to open a file with SDB report and getting its full path
    With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
    .Show
    fullpath = .SelectedItems.Item(1)
    If InStr(fullpath, ".xls") = 0 Then
            Exit Sub
        End If
    End With

'to open a chosen file
    Set wb = Workbooks.Open(fullpath, ReadOnly:=True)
    Set ws = wb.Worksheets("dp_log")
    
    lRowTemp = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'since SDB report is carried by not responsible person need to do some changes
    For i = 2 To lRowTemp
        Set cellCompared = ws.Cells(i, 2)
        cellCompared.Value = Replace(Replace(Replace(cellCompared.Value, "DDP", "DP"), "-", "."), " ", "")
    Next i

'Formating Status column from lowercase to upper
    For i = 2 To lRowTemp
        Set cellCompared = ws.Cells(i, 4)
        cellCompared.Value = Replace(Replace(cellCompared.Value, "code", "Code"), "Code1", "Code 1")
    Next i
    
'errors ingore
    ws.Range(ws.Cells(2, 3), ws.Cells(lRow, 3)).Value = ws.Range(ws.Cells(2, 3), ws.Cells(lRow, 3)).Value
    
'to create arrays with module, dp, revision and status information
    
    ReDim sdbM(1 To lRowTemp)
    ReDim sdbDP(1 To lRowTemp)
    ReDim sdbRev(1 To lRowTemp)
    ReDim sdbStatus(1 To lRowTemp)
    
    sdbM = ws.Range(ws.Cells(2, 1), ws.Cells(lRowTemp, 1)).Value
    sdbDP = ws.Range(ws.Cells(2, 2), ws.Cells(lRowTemp, 2)).Value
    sdbRev = ws.Range(ws.Cells(2, 3), ws.Cells(lRowTemp, 3)).Value
    sdbStatus = ws.Range(ws.Cells(2, 4), ws.Cells(lRowTemp, 4)).Value
    
'close SDB report
    wb.Close False
    
'to fill data as needed
    For i = 2 To lRow
    
    'to declare the base cell
        Set cellChecked = mainWs.Cells(i, 1)
    'to declare module and DP concat
        mDPbase = cellChecked.Value
        
    'term
        term = 1
        
        'to start compare module and DP concat with SDB report
        For j = 1 To lRowTemp - 1
            
        'to declare compared Module and DP
            mDPcomp = Mid(sdbM(j, 1), 1, 1) & Mid(sdbM(j, 1), 5, 1) & Mid(sdbM(j, 1), 9, 1) & sdbDP(j, 1)
            
        'if matches copy it and change cell format
            If mDPbase = mDPcomp And term = sdbRev(j, 1) Then
                cellChecked.Offset(0, term).Value = sdbStatus(j, 1)
                
                'to paint the cells based on Code number
                Select Case sdbStatus(j, 1)
                    Case "Code 1"
                        cellChecked.Offset(0, term).Interior.Color = RGB(255, 51, 0)
                    Case "Code 2"
                        cellChecked.Offset(0, term).Interior.Color = RGB(0, 204, 153)
                    Case "Code 3"
                        cellChecked.Offset(0, term).Interior.Color = RGB(51, 153, 102)
                    Case "Submitted to CTR"
                        cellChecked.Offset(0, term).Interior.Color = RGB(102, 178, 255)
                        cellChecked.Offset(0, term).Value = "Under review"
                        cellChecked.Offset(0, 11).Value = "Under review"
                    Case "Re-Submitted to CTR"
                        cellChecked.Offset(0, term).Interior.Color = RGB(102, 178, 255)
                        cellChecked.Offset(0, term).Value = "Under review"
                        cellChecked.Offset(0, 11).Value = "Under review"
                    End Select
                
                'to insert last revision
                cellChecked.Offset(0, 11).Value = sdbStatus(j, 1)
                
                term = term + 1
            End If
            
        Next j
        
        'last revision column filling in a case when it's empty
        If cellChecked.Offset(0, 1).Value = "" Then
            cellChecked.Offset(0, 11).Value = "Not Submitted"
        End If
        
        'to paint the cells based on Code number
        Select Case cellChecked.Offset(0, 11).Value
            Case "Code 1"
                cellChecked.Offset(0, 11).Interior.Color = RGB(255, 51, 0)
            Case "Code 2"
                cellChecked.Offset(0, 11).Interior.Color = RGB(0, 204, 153)
            Case "Code 3"
                cellChecked.Offset(0, 11).Interior.Color = RGB(51, 153, 102)
            Case "Submitted to CTR"
                cellChecked.Offset(0, 11).Value = "Under review"
                cellChecked.Offset(0, 11).Interior.Color = RGB(102, 178, 255)
            Case "Re-Submitted to CTR"
                cellChecked.Offset(0, 11).Value = "Under review"
                cellChecked.Offset(0, 11).Interior.Color = RGB(102, 178, 255)
            Case "Not Submitted"
                cellChecked.Offset(0, 11).Interior.Color = RGB(255, 255, 153)
        End Select
    Next i
    
    mainWs.Protect AllowFiltering:=True
    Application.Calculation = xlCalculationAutomatic
    

End Sub

