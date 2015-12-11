Const VK_CAPITAL = 20
Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Sub pressWord(str As String)

    Dim c As String
    Dim iter As Integer
    
   Call delay(0.5)
    
    For iter = 1 To Len(str)
        
        c = Mid(str, iter, 1)
        
        If Asc(c) < 127 Then
            Select Case Asc(c)
                Case 65 To 90 'ASCII Code for A to Z
                    c = UCase(c)
                    keybd_event VK_CAPITAL, 0, 0, 0
                    keybd_event VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0
                    keybd_event Asc(c), 0, 0, 0
                    keybd_event Asc(c), 0, KEYEVENTF_KEYUP, 0
                    keybd_event VK_CAPITAL, 0, 0, 0
                    keybd_event VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0
                Case Else
                    c = UCase(c)
                    keybd_event Asc(c), 0, 0, 0
                    keybd_event Asc(c), 0, KEYEVENTF_KEYUP, 0
                End Select
        End If
        Call delay(0.5)
    Next iter

End Sub


Private Sub delay(seconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", seconds, Now())
Do While Now() < endTime
        DoEvents
    Loop
End Sub