Sub significantDigits()
    'improved from
    'http://www.spreadsheet-validierung.de/excel-signifikante-stellen/
    
    Dim digits, target As Integer
    Dim tmp, unitsign As String
    'Dim cel, selectedRange As Range
    Dim start As Date
    svc = createUnoService( "com.sun.star.sheet.FunctionAccess" )
	DIM oNumberFormats AS OBJECT
	oNumberFormats = ThisComponent.getNumberFormats()

    target = 3 'target amount of significant decimal digits
    selectedRange = ThisComponent.CurrentSelection

    'From https://ask.libreoffice.org/t/calc-loop-through-all-cells-in-the-current-selection/49185/5
    If selectedRange.supportsService("com.sun.star.sheet.SheetCellRanges") Then
	  rgs = selectedRange
	Else
  	If NOT selectedRange.supportsService("com.sun.star.sheet.SheetCellRange") Then Exit Sub
  		rgs = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
  		rgs.addRangeAddress(selectedRange.RangeAddress, False)
	End If

    start = Now
    For Each rg In rgs
    For i = 0 To rg.Rows.getCount() - 1             
        For j = 0 To rg.Columns.getCount() - 1
            Set cel = rg.getCellByPosition( j, i )
        If DateDiff("s", start, Now) > 5 Then
            MsgBox ("Timeout 5 s reacched. Make a smaller selection.")
            Exit Sub
        End If
        tmp = cel.Value

		nfCode = oNumberFormats.getByKey(cel.NumberFormat).FormatString
        If IsNumeric(tmp) And Not IsEmpty(tmp) Then
            If InStr(nfCode, "%") Then
                unitsign = " %"
                tmp = tmp * 100
            ElseIf InStr(nfCode, "€") Then
                unitsign = " €"
                Print "EUR"
            ElseIf InStr(nfCode, "$") Then
                unitsign = " $"
            Else
                unitsign = ""
            End If
            If tmp = 0 Then
                digits = 0
            Else
                digits = -Int(Log(Abs(tmp)) / Log(10#)) + target - 1
                digits = svc.callFunction("MAX",Array(0, digits))
            End If
            'remove trailing zeroes:
            Do While True
                If digits > 0 Then
                	'print oNumberFormats.queryKey(stNumberFormat, aLocale, FALSE)
                    If svc.callFunction("Round",Array(tmp, svc.callFunction("MAX",Array(digits + 10, 22)))) = svc.callFunction("Round",Array(tmp, digits - 1)) Then
                        digits = digits - 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            
            If digits = 0 Then
	        	cel.NumberFormat = CellSetNumberFormat("0" & unitsign, oNumberFormats)
            Else
            	cel.NumberFormat = CellSetNumberFormat("0," & String(CLng(digits), "0") & unitsign, oNumberFormats)
            End If
        End If
    Next
    Next 
    Next rg
    
End Sub

FUNCTION CellSetNumberFormat(stNumberFormat AS STRING, oNumberFormats AS OBJECT) AS LONG
	'https://ask.libreoffice.org/t/basic-numberformat-codes/88899
	DIM aLocale	AS NEW com.sun.star.lang.Locale
	DIM loFormatKey	AS LONG
	loFormatKey = oNumberFormats.queryKey(stNumberFormat, aLocale, FALSE)
	IF loFormatKey = -1 THEN loFormatKey = oNumberFormats.addNew(stNumberFormat, aLocale)
	CellSetNumberFormat = loFormatKey
End Function

