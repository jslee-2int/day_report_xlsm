Attribute VB_Name = "Module1"
Sub add_next_week()
    Dim ws, week_num, exist_sheet As Integer
    Dim ws_name
    
    Dim sht As Worksheet
    Dim i As Integer
    Dim month, day, weekday, start_day
    Dim week_str
    week_str = Array("��", "ȭ", "��", "��", "��")
    
    exist_sheet = 0
    
    ws = Worksheets.Count
    week_num = DatePart("ww", Now)
    week_num = CInt(week_num) - 2
    
    Sheets(1).Copy after:=Sheets(ws)
    week_num = CInt(week_num) + 1
    ws_name = "W" & week_num
    ws = ws + 1
    
    For Each sht In Worksheets
        If sht.name = ws_name Then
            exist_sheet = 1
            MsgBox "�̹��� ��Ʈ�� ���� �����մϴ�."
        End If
    Next sht
    
    If exist_sheet = 0 Then
        Sheets(ws).name = ws_name
    Else
        Application.DisplayAlerts = False
        Sheets(ws).Delete
    End If
    
    Cells(2, 2).Value = ws_name + " �������� �� ��ȹ"
    
    month = DatePart("m", Now)
    day = DatePart("d", Now)
    weekday = DatePart("w", Now)
        
    day = -CInt(weekday) + 2
        
    For i = 0 To 4
        start_day = DateAdd("d", day + i, Now)
        Cells(6 + i, 2).Value = week_str(i) & Chr(13) & Chr(10) & "(" & DatePart("m", start_day) & "�� " & DatePart("d", start_day) & "��" & ")"
    Next
    'MsgBox DateAdd("d", day, Now)
    
End Sub
