Sub Categories():

Dim categories() As string
Dim i As Integer

For i = 2 To 4115
    categories = Split(Cells(i, 14), "/")
    Cells(i, 17).Value = categories(0)
    Cells(i, 18).Value = categories(1)
Next i

End Sub