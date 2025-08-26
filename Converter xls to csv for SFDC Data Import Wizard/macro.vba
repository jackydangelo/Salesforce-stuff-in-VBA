'Macro generate csv for data import wizard D'angelo 04/05/2020
'Optimization work 5 years later since first use 21/08/2025
Sub Main()
    CreateFile
End Sub

Sub QuoteCommaExport(ws As Worksheet, percorso As String)
    Dim fsT As Object
    Dim riga As Long, colonna As Long
    Dim ultimaRiga As Long, ultimaColonna As Long
    Dim lineaCSV As String
    Dim valori() As String

    Set fsT = CreateObject("ADODB.Stream")
    With fsT
        .Type = 2
        .Charset = "UTF-8"
        .Open
    End With

    With ws
        ultimaRiga = .Cells(.Rows.Count, 1).End(xlUp).Row
        ultimaColonna = .Cells(1, .Columns.Count).End(xlToLeft).Column

        For riga = 1 To ultimaRiga
            ReDim valori(1 To ultimaColonna)
            For colonna = 1 To ultimaColonna
                valori(colonna) = """" & Replace(.Cells(riga, colonna).Text, """", """""") & """"
            Next colonna
            lineaCSV = Join(valori, ",")
            fsT.WriteText lineaCSV & vbCrLf
        Next riga
    End With

    fsT.SaveToFile percorso, 2
    fsT.Close
    Set fsT = Nothing
End Sub

Sub CreateFile()
    Dim percorso As String
    Dim sName As String: sName = "esempio.csv"
    Dim fd As FileDialog
    Dim ws As Worksheet

    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    Set ws = ActiveSheet

    With fd
        .Title = "Salva il file CSV"
        .InitialFileName = sName
        .FilterIndex = GetFilterIndex
        If .Show = -1 Then
            percorso = .SelectedItems(1)
            QuoteCommaExport ws, percorso
            MsgBox "Esportazione completata! Il file si trova nella cartella selezionata.", vbInformation
        End If
    End With
End Sub

Function GetFilterIndex() As Integer
    Dim versione As Integer
    versione = Val(Split(Application.Version, ".")(0))

    Select Case versione
        Case 14: GetFilterIndex = 15 ' Excel 2010
        Case 15: GetFilterIndex = 1  ' Excel 2013
        Case 16: GetFilterIndex = 16 ' Excel 2016
        Case Else: GetFilterIndex = 1
    End Select
End Function
