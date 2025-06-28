Sub FilterByRiskLevel(riskLevel As String)

    Dim riskWS As Worksheet: Set riskWS = Worksheets("risk_list")
    Dim transWS As Worksheet: Set transWS = Worksheets("transactions")

    Dim countryList As String
    Dim i As Long, lastRow As Long

    ' Zbierz kraje o wybranym poziomie ryzyka
    With riskWS
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            If .Cells(i, 2).Value = riskLevel Then
                countryList = countryList & "," & .Cells(i, 1).Value
            End If
        Next i
    End With

    If countryList = "" Then
        MsgBox "Brak krajów o ryzyku: " & riskLevel, vbExclamation
        Exit Sub
    End If

    countryList = Mid(countryList, 2) ' usuń pierwszy przecinek

    ' Filtruj transakcje wg kraju (Country = kolumna 6)
    With transWS
        .AutoFilterMode = False
        .Range("A1").AutoFilter Field:=6, Criteria1:=Split(countryList, ","), Operator:=xlFilterValues
    End With

End Sub



Sub FilterHigh()
    FilterByRiskLevel "High"
End Sub

Sub FilterMedium()
    FilterByRiskLevel "Medium"
End Sub

Sub FilterLow()
    FilterByRiskLevel "Low"
End Sub

Sub ResetFilter()
    Dim ws As Worksheet
    Set ws = Worksheets("transactions")

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(1) ' lub podaj nazwę, np. ws.ListObjects("TransactionsTable")

    If tbl.AutoFilter.FilterMode Then
        tbl.AutoFilter.ShowAllData
    End If
End Sub

