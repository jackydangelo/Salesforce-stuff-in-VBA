Sub MappingValue()
Dim i As Integer

i = 2
Do While Not IsEmpty(Worksheets("Sheet1").Cells(i, 4))
Select Case Worksheets("Sheet1").Cells(i, 4).Value

Case "Libero"
Worksheets("Sheet1").Cells(i, 4) = "L"
Case "Salvaguardia"
Worksheets("Sheet1").Cells(i, 4) = "S"
Case "Tutela"
Worksheets("Sheet1").Cells(i, 4) = "T"

End Select
i = i + 1
Loop
End Sub
