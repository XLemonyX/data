Sub ClearMAKRO()
    Dim ws As Worksheet
    Set ws = Sheets("MAKRO")

    ' 🧼 Usuń wszystkie kontrolki ActiveX i Form Control (tylko przyciski)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then
            If shp.Name Like "btn*" Then
                shp.Delete
            End If
        End If
    Next shp

    ' 🧹 Wyczyść wiersze od 2 w dół
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ws.Rows("2:" & lastRow).ClearContents
    End If

    MsgBox "Zakładka MAKRO została wyczyszczona – dane i wszystkie przyciski usunięte 🧽", vbInformation
End Sub
