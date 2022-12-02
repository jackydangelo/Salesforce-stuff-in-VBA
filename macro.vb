'Macro generate csv for data import wizard D'angelo 04/05/2020

Sub main()
    
    'Here we can user extra n 1 "MappingValue.vb"
    CreateFile
    
End Sub


Sub QuoteCommaExport(percorso)

   ' Dimension all variables.
   Dim DestFile As String
   Dim FileNum As Integer
   Dim ColumnCount As Integer
   Dim RowCount As Integer
   Dim fsT As Object
   
   Cells.Select
   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "UTF-8" 'Specify charset For the source text data.
   fsT.Open 'Open the stream And write binary data To the object
    
   ' file destination string recovery
   DestFile = percorso

   ' select all cells
   ActiveCell.CurrentRegion.Select
   ' Loop per each row
   For RowCount = 1 To Selection.Rows.Count

      ' Loop per each column
      For ColumnCount = 1 To Selection.Columns.Count
      fsT.WriteText """" & Selection.Cells(RowCount, ColumnCount).Text & """"
        ' Verify if i am in the last column
         If ColumnCount = Selection.Columns.Count And RowCount <> Selection.Rows.Count And RowCount = 1 Then
            fsT.WriteText "" & vbCrLf
        ElseIf ColumnCount = Selection.Columns.Count And RowCount <> Selection.Rows.Count And RowCount <> 1 Then
             fsT.WriteText "," & vbCrLf
        ElseIf ColumnCount = Selection.Columns.Count And RowCount = Selection.Rows.Count Then
            fsT.WriteText ""
         Else
            fsT.WriteText ","
         End If
      Next ColumnCount
   Next RowCount
   fsT.SaveToFile DestFile, 2 'save file
End Sub


Sub CreateFile()
' Sub for the creation of savedialog and generation of csv
Dim percorso As String
Dim Cartella As String
Dim sName As String
sName = "esempio"

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogSaveAs)
Dim CartellaSelezionata As Variant
With fd
fd.Title = "Seleziona la cartella"

fd.InitialFileName = sName
fd.FilterIndex = versionExcel

If .Show = -1 Then
For Each CartellaSelezionata In .SelectedItems

miaCartella = CartellaSelezionata

Next
Else
Exit Sub
End If
percorso = fd.SelectedItems(1)
End With

QuoteCommaExport percorso
MsgBox ("Fatto! Il file generato è presente nella cartella selezionata"            
End Sub
            
            
 Function versionExcel() As Integer
    ' function for find version of excel used
    ' with this parameter we can suggest the correct extension (csv)
    Dim xlApp As New Excel.Application

    Select Case Val(Mid(xlApp.Version, 1, _
        InStr(1, xlApp.Version, ".") - 1))
        Case 14 'index of "Excel 2010"
            versionExcel = 15 'index csv for "Excel 2010"
        Case 15 'index of "Excel 2013"
            versionExcel = 1 'index csv for "Excel 2013"
        Case 16 'index of "Excel 2016"
            versionExcel = 16 'index csv for "Excel 2016"
        Case Else
            versionExcel = 1 'default
    End Select
    
    Set xlApp = Nothing
End Function
