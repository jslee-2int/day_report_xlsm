Attribute VB_Name = "Module3"
Option Explicit

Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal _
 hwnd As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
 "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName _
 As Any) As Long
 
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'키입력을 위한 선언하기
Const KEYEVENTF_EXTENDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Private Declare PtrSafe Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


Function setClip(str)

Dim obj As Object
Set obj = CreateObject( _
                 "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
obj.setText str
obj.PutInClipboard
'SetCB = True

End Function

Function getClip$()

    Dim obj As Object
    Set obj = CreateObject( _
                 "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.GetFromClipboard
    getClip = obj.GetText

End Function

Sub web_control()

    'Dim IE_ctrl As InternetExplorer
    Dim IE_ctrl As Object
    Set IE_ctrl = CreateObject("InternetExplorer.Application")
    Dim HTMLDoc As IHTMLDocument
    Dim input_Data As IHTMLElement
    Dim URL As String
    
    Set IE_ctrl = New InternetExplorer
    IE_ctrl.Silent = True
    IE_ctrl.Visible = True
    
    Dim THandle As Long
    Dim Alllinks, Hyperlink
 
    THandle = FindWindow("IEFrame", vbEmpty)
    
    If THandle = 0 Then
     MsgBox "Could not find window.", vbOKOnly
    Else
     BringWindowToTop THandle
    End If
    
    Sleep (1000)
    'Application.SendKeys ("{Enter}")
    
    URL = "https://office.hiworks.com/your_domain/home/logout" '로그아웃 URL
    IE_ctrl.navigate URL
    Sleep (1000)
    
    URL = "https://office.hiworks.com/your_domain/bbs/board/board_write/modify/4321/123" '본인 게시물 주소
    IE_ctrl.navigate URL
    
    Sleep (1100)
    Application.SendKeys ("{Enter}")
    
    Do Until IE_ctrl.readyState = 4
       DoEvents
    Loop
    
    Set HTMLDoc = IE_ctrl.document

    Set input_Data = HTMLDoc.getElementById("office_id")
    input_Data.Value = "Your ID" 'ID 입력
    
    Set input_Data = HTMLDoc.getElementById("office_passwd")
    input_Data.Value = "Your PW" '암호 입력
    
    Set input_Data = HTMLDoc.getElementsByClassName("int_jogin").Item
    input_Data.Click
    
    Sleep (1100)
    
    Set input_Data = HTMLDoc.getElementsByClassName("icon file_delete").Item
    input_Data.Click
    
    Sleep (1100)
    
    Set input_Data = HTMLDoc.getElementsByClassName("weakblue unuserble_9_8").Item
    input_Data.Click
    
    '파일 경로 수정
    Sleep (1100)
    setClip ("D:\Documents\일일업무보고_양식.xlsm")
    
    Dim result
    result = getClip()
    
    Application.SendKeys ("^v")
    
    Sleep (1100)
    Application.SendKeys ("{Enter}")
    
    Sleep (1000)
    
    Set Alllinks = HTMLDoc.getElementsByClassName("detail_select")
    Set Alllinks = HTMLDoc.getElementsByTagName("A")
    
    For Each Hyperlink In Alllinks
        If Hyperlink.innerText = "확인" Then
            Hyperlink.Click
        End If
    Next Hyperlink
    
    Sleep (1100)
    Application.SendKeys ("{Enter}")
    Sleep (1100)
    
    IE_ctrl.Quit
    Set IE_ctrl = Nothing
    
    Application.SendKeys ("{NUMLOCK}")
    
End Sub



