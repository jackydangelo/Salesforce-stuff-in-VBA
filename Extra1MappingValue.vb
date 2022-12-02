Sub MappingValue()
Dim i As Integer

i = 2
Do While Not IsEmpty(Worksheets("Account").Cells(i, 4))
Select Case Worksheets("Account").Cells(i, 4).Value

Case "Libero"
Worksheets("Account").Cells(i, 4) = "L"
Case "Salvaguardia"
Worksheets("Account").Cells(i, 4) = "S"
Case "Tutela"
Worksheets("Account").Cells(i, 4) = "T"

End Select
i = i + 1
Loop
End Sub
