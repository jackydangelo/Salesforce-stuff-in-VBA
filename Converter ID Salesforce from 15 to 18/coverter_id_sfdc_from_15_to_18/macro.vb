Sub main()

    ConvertSalesforceIDsAuto

End Sub

Sub ConvertSalesforceIDsAuto()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim id15 As String
    Dim id18 As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        id15 = Trim(ws.Cells(i, 1).Value)
        id18 = ConvertSalesforceID(id15)
        ws.Cells(i, 2).Value = id18
    Next i
End Sub

Function ConvertSalesforceID(id15 As String) As String
    Dim suffix As String
    Dim i As Integer, j As Integer
    Dim block As String
    Dim bitString As String
    Dim bitValue As Integer
    Dim lookup As String
    
    lookup = "ABCDEFGHIJKLMNOPQRSTUVWXYZ012345"
    suffix = ""
    
    If Len(id15) <> 15 Then
        ConvertSalesforceID = "ID non valido"
        Exit Function
    End If
    
    For i = 0 To 2
        block = Mid(id15, i * 5 + 1, 5)
        bitString = ""
        
        ' Costruzione corretta della stringa binaria (ordine normale)
        For j = 1 To 5
            If Mid(block, j, 1) Like "[A-Z]" Then
                bitString = bitString & "1"
            Else
                bitString = bitString & "0"
            End If
        Next j
        
        bitValue = 0
        For j = 0 To 4
            If Mid(bitString, j + 1, 1) = "1" Then
                bitValue = bitValue + 2 ^ j
            End If
        Next j
        
        suffix = suffix & Mid(lookup, bitValue + 1, 1)
    Next i
    
    ConvertSalesforceID = id15 & suffix
End Function

