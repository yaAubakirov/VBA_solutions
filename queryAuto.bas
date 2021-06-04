Attribute VB_Name = "Module1"
Sub queryAuto()

Dim lRow As Long
Dim lCol As Long
Dim cellChecked As Range
Dim module As String
Dim dpCount As Long
Dim dpNumber As String
Dim splitString() As String

'find last row index
lRow = Cells(Rows.Count, 1).End(xlUp).Row

'filling General column
For i = 3 To lRow

    If InStr(Cells(i, 1), "General") > 0 Then
        Cells(i, 1).Offset(0, 5).Value = "Yes"
    Else:
        Cells(i, 1).Offset(0, 5).Value = "No"
    End If

Next i

'find last column index
lCol = Cells(3, Columns.Count).End(xlToLeft).Column

'deleting shitty stuff as comas and spaces
For i = 3 To lRow
    Set cellChecked = Cells(i, lCol)
    
    cellChecked.Value = Replace(cellChecked, ",", "")
    cellChecked.Value = Replace(cellChecked, "DP ", "DP")

Next i

'filling in required format
For i = 3 To lRow
    'cell assign
    Set cellChecked = Cells(i, lCol)
    
    'spliting the DPs
    splitString = Split(cellChecked, " ")
    
    'module name
    module = Mid(Cells(i, 1), 10, 1) & Mid(Cells(i, 1), 13, 1) & Mid(Cells(i, 1), 16, 1)
    
    'how many dps in TQ
    dpCount = Len(Cells(i, 7)) - Len(Replace(Cells(i, 7), " ", "")) + 1
    
    'checks if it is general
    If Cells(i, 6) = "No" Then
        
        'checks if dp number is more than 1
        If dpCount > 1 Then
    
            j = 1
            
            'concatenate the module with splited DP using its index
            Do Until j = dpCount + 1
        
                cellChecked.Offset(0, j).Value = module & splitString(j - 1)
                j = j + 1
            
            Loop
        
        'same for single dp concat with cell value
        Else:
    
            cellChecked.Offset(0, 1).Value = module & cellChecked.Value
        
        End If
        
    End If
        
Next i

MsgBox "Perpecto!"

End Sub
