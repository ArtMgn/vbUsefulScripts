Sub UnhideAllSheets()
    Dim ws As Worksheet
 
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
 
End Sub


Public Function optimize(ByVal opt As Boolean) As Boolean
    Select Case opt
        Case True
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
            Application.EnableEvents = False
        Case False
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        End Select
        optimize = True
End Function