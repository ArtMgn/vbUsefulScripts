Option Explicit


Public Sub getAllTable()

    Dim max As Integer
    max = findMaxColumn
    Debug.Print max
    
    Dim var As Variant
    Dim lastR As Double
    
    lastR = getLastRow("A1")
    
    var = Range("A1", Cells(lastR, max))

    call imptab(var)
    
    
End Sub


Public Function findMaxColumn() As Integer

    Dim max As Integer
    Dim c As Variant
    Dim i As Integer
    
    
    For Each c In Range("A1", Range("A1").End(xlDown).Address)
        Do While c.Offset(0, i) <> ""
            i = i + 1
        Loop
        If i > max Then max = i
        i = 0
    Next c
    
    findMaxColumn = max
    
End Function


Public Function findMaxColumn2() as Integer

    Dim max as Integer
    Dim lastR as Double
    Dim rows as String

    lastR = getLastRow("A1")
    rows = "1:" & lastR
    max = Evaluate("=MAX((" & rows & "<>"""")*COLUMN(" & rows & "))")
 
    findMaxColumn2 = max

End Function


Public Function getLastRow(ByVal str As String) As Long
    getLastRow = Range(str).End(xlDown).row
End Function


Sub imptab(x As Variant)

    Dim ligne As Long
    Dim colonne  As Long
    Dim n As Variant
    Dim maligne As String

    For ligne = LBound(x, 1) To UBound(x, 1)
        maligne = ""
        For colonne = LBound(x, 2) To UBound(x, 2)
            maligne = maligne & x(ligne, colonne) & " "
        Next colonne
        Debug.Print Left(maligne, Len(maligne) - 1)
    Next ligne

End Sub