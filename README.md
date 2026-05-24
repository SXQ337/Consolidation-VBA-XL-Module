# Consolidation-VBA-XL-Module
This code is used to extract the required data from the total data.

```vba
Public Sub First_Question_Ans_2nd()
Dim i As Long
Dim rowindex As Long
Worksheets(1).Select

Range("C2").End(xlDown).Select
rowindex = ActiveCell.Row

For i = 2 To rowindex
    If Cells(i, 3).Value = "Life" And _
        Cells(i, 4).Value > 8000 And _
        Cells(i, 5).Value = "Active" Then
    
        With Range("A" & i & ":E" & i)
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
        End With
    
    ElseIf Cells(i, 4).Value > 5000 And _
            Cells(i, 5).Value = "Lapsed" Then
            
            With Range("A" & i & ":E" & i)
                .Interior.Color = RGB(255, 0, 0)
                .Font.Bold = True
        End With
    
    Else:
        With Range("A" & i & ":E" & i)
            .Interior.Color = RGB(204, 255, 204)
            .Font.Bold = True
        End With
    End If
Next i

Worksheets(1).Name = "Insurance Consolidation"
Worksheets(2).Name = "High Value Cost"
Worksheets(3).Name = "Lapsed Insurance"

End Sub

Public Sub HighValueCost_Demo()
Worksheets("High Value Cost").Select
Range("A1").Select
Selection.Value = "High Value Cost Template"
Range("A2:E2").Select
Selection.Value = Array("Insurance ID", "Clients Name", "Insurance Category", _
"Insurance Claims", "Status")
Selection.Interior.Color = RGB(255, 0, 255)
With Selection.Font
    .Size = 20
    .Bold = True
    .Name = "Times New Roman"
End With
Columns.AutoFit

Worksheets(1).Select
Range("C2").Select
Selection.End(xlDown).Select
rowindex = ActiveCell.Row

For i = 2 To rowindex
    If Cells(i, 3).Value = "Life" And _
        Cells(i, 4).Value > 8000 And _
        Cells(i, 5).Value = "Active" Then
        
        Rows(i).Copy
        Sheets(2).Select
        Rows(i).PasteSpecial
        Sheets(1).Select
    End If
Next i
Sheets(2).Select
Call NoBlank
End Sub

Public Sub LapsedIns()
Worksheets("Lapsed Insurance").Select
Range("A1").Select
Selection.Value = "Lapsed Template"
Range("A2:E2").Select
Selection.Value = Array("Insurance ID", "Clients Name", "Insurance Category", _
"Insurance Claims", "Status")
Selection.Interior.Color = RGB(255, 0, 255)
With Selection.Font
    .Size = 20
    .Bold = True
    .Name = "Times New Roman"
End With
Columns.AutoFit

Worksheets(1).Select
Range("C2").Select
Selection.End(xlDown).Select
rowindex = ActiveCell.Row

For i = 2 To rowindex
     If Cells(i, 4).Value > 5000 And _
        Cells(i, 5).Value = "Lapsed" Then
        
        Rows(i).Copy
        Sheets(3).Select
        Rows(i).PasteSpecial
        Sheets(1).Select
    End If
Next i
Sheets(3).Select
Call NoBlank
End Sub

Sub NoBlank()
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("A1").Select
End Sub

Public Sub Clear()
    Range("A2", Range("E2").End(xlDown)).Select
    Selection.ClearFormats
    Range("A1", Range("n500")).Clear
    Range("A1", Range("n500")).ClearFormats
End Sub
```
