Sub significantDigits()
    'improved from
    'http://www.spreadsheet-validierung.de/excel-signifikante-stellen
    Dim digits, target As Integer
    Dim tmp, unitsign As String
    Dim cel, selectedRange As Range
    Set selectedRange = Application.Selection
    target = 3 'target amount of significant decimal digits
    Dim start As Date
    start = Now
    For Each cel In selectedRange.Cells
        If DateDiff("s", start, Now) >= 5 Then
            MsgBox ("Timeout 5 s reached. Make a smaller selection.")
            Exit Sub
        End If
        tmp = cel.Value
        If IsNumeric(tmp) And Not IsEmpty(tmp) Then
            If InStr(cel.NumberFormat, "%") Then
                unitsign = " %"
                tmp = tmp * 100
            ElseIf InStr(cel.NumberFormat, "€") Then
                unitsign = " €"
            ElseIf InStr(cel.NumberFormat, "$") Then
                unitsign = " $"
            Else
                unitsign = ""
            End If
            If tmp = 0 Then
                digits = target - 1
            Else
                digits = -Int(Log(tmp) / Log(10#)) + target - 1
                digits = WorksheetFunction.Max(0, digits)
            End If
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
                cel.NumberFormat = "0" & unitsign
            Else
                cel.NumberFormat = "0." & String(CLng(digits), "0") & unitsign
            End If
        End If
    Next cel
End Sub
