Attribute VB_Name = "significantDigits"
Sub significantDigits()
'improved from
'http://www.spreadsheet-validierung.de/excel-signifikante-stellen
Dim i, digits, rows, cols, nRow, target As Integer
Dim tmp, percentsign As String
target = 3 'target amount of significant decimal digits
rows = Selection.Row
cols = Selection.Column
For i = 1 To Selection.Count
    nRow = rows + i - 1
    tmp = Cells(nRow, cols)
    If InStr(Cells(nRow, cols).NumberFormat, "%") Then
        percentsign = "%"
        tmp = tmp * 100
    Else
        percentsign = ""
    End If
    digits = -Log(tmp) / Log(10#) + target - 1
    digits = WorksheetFunction.Max(0, digits)
    'remove trailing zeroes:
    Do While True
        If digits > 0 Then
            If tmp = Round(tmp, digits - 1) Then
                digits = digits - 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    '
    If digits = 0 Then
        Cells(nRow, cols).NumberFormat = "0" & percentsign
    Else
        Cells(nRow, cols).NumberFormat = "0." & String(CLng(digits), "0") & percentsign
    End If
Next i
End Sub


