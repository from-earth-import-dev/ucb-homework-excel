Sub AverageDonation():

Dim pledged As Long
Dim backers As Long

Dim i As Integer

For i = 2 To 4115
    backers = Cells(i, 12)
    pledged = Cells(i, 5)
    percent_funded = Cells(i, 16)

    Cells(i, 16).Value = Round((pledged / backers), 2)
Next i

End Sub