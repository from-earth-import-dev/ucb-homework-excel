Sub PercentFunded():

Dim goal As Long
Dim pledged As Long

Dim i As Integer

For i = 2 to 4115
    goal = Cells(i, 4)
    pledged = Cells(i, 5)
    percent_funded = Cells(i, 15)

    Cells(i, 15).Value = (pledged / goal) * 100
Next i

End Sub
