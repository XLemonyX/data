On Error Resume Next
Dim rawVal As Variant
rawVal = wsSource.Cells(i, 14).Value
clientName = "UNKNOWN" ' domyślna wartość na wypadek błędu

If Not IsError(rawVal) Then
    If VarType(rawVal) = vbString Then
        clientName = Trim(rawVal)
    ElseIf VarType(rawVal) = vbDouble Or VarType(rawVal) = vbInteger Or VarType(rawVal) = vbLong Then
        clientName = Trim(CStr(rawVal))
    ElseIf VarType(rawVal) = vbNull Or VarType(rawVal) = vbEmpty Then
        clientName = "EMPTY"
    Else
        clientName = "UNEXPECTED"
    End If
Else
    clientName = "ERROR"
End If
On Error GoTo 0




' Tytu? oczyszczony
If InStr(UCase(title), "VIPPS") > 0 Or InStr(UCase(title), "TRANSREF") > 0 Then
cleanTitle = ""
Else
cleanTitle = NormalizeSpaces(title)
End If

' Sumowanie
If Not summary.exists(clientName) Then
summary.Add clientName, Array(inAmount, outAmount)
cleanTitles.Add clientName, cleanTitle
Else
Dim temp(): temp = summary(clientName)
temp(0) = temp(0) + inAmount
temp(1) = temp(1) + outAmount
summary(clientName) = temp

If cleanTitle <> "" Then
If cleanTitles(clientName) = "" Then
cleanTitles(clientName) = cleanTitle
Else
cleanTitles(clientName) = cleanTitles(clientName) & "; " & cleanTitle
End If
End If
End If
End If
Next i

' Wypisanie do arkusza
Dim rowIndex As Long: rowIndex = 2
Dim totalIn As Double: totalIn = 0
Dim totalOut As Double: totalOut = 0
Dim key As Variant
For Each key In summary.Keys
Dim values(): values = summary(key)
wsReport.Cells(rowIndex, 1).Value = key
wsReport.Cells(rowIndex, 2).Value = values(0)
wsReport.Cells(rowIndex, 3).Value = values(1)
wsReport.Cells(rowIndex, 4).Value = cleanTitles(key)
totalIn = totalIn + values(0)
totalOut = totalOut + values(1)
rowIndex = rowIndex + 1
Next key

' Formatowanie wynikďż˝w
With wsReport
.Columns("A:D").AutoFit
.Range("A1:D1").Font.Bold = True
.Range("A1:D1").Interior.Color = RGB(0, 102, 204)
.Range("A1:D1").Font.Color = RGB(255, 255, 255)
.Range("A2:D" & rowIndex - 1).Borders.Color = RGB(180, 210, 240)
.Range("B2:C" & rowIndex - 1).NumberFormat = "# ##0"
.Range("A1:D" & rowIndex - 1).AutoFilter
.Range("A2:D" & rowIndex - 1).Sort Key1:=.Range("B2"), Order1:=xlDescending, Header:=xlNo
.Range("B2:C" & rowIndex - 1).Font.Bold = True
End With


' Podsumowanie ogďż˝lne obok tabeli
Dim summaryStart As Range
Set summaryStart = wsReport.Cells(2, 6) ' Kolumna F

With summaryStart
.Value = "Total people:"
.Offset(0, 1).Value = summary.Count
.Offset(1, 0).Value = "Total received:"
.Offset(1, 1).Value = totalIn
.Offset(2, 0).Value = "Total sent:"
.Offset(2, 1).Value = totalOut

.EntireColumn.AutoFit
.Resize(3, 2).Font.Bold = True
.Resize(3, 2).Borders.Weight = xlThin
.Resize(3, 2).Interior.Color = RGB(230, 240, 255)

' Formatowanie liczby osďż˝b jako liczby ca?kowitej
.Offset(0, 1).NumberFormat = "0"

' Formatowanie sum bez miejsc po przecinku i bez waluty
.Offset(1, 1).NumberFormat = "# ##0"
.Offset(2, 1).NumberFormat = "# ##0"
End With

End Sub



