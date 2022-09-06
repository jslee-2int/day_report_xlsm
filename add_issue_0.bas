Attribute VB_Name = "Module2"
Sub add_issue_0()
    Dim row_num As Integer
    Dim currentSelection As Range
    Dim text As String
    
    row_num = ActiveCell.Row + 1
    Rows(row_num).Insert
    
    Set currentSelection = Selection.Cells()
    
    text = currentSelection.Offset(-1, 0).Value
    'If text = "" Then
    Range(currentSelection.Offset(0, -5), currentSelection.Offset(1, -5)).Select
    Range(currentSelection.Offset(0, -5), currentSelection.Offset(1, -5)).Merge
    'End If
    currentSelection.Offset(0, -4).Select
End Sub

