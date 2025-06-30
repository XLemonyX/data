    Sub ClearMAKRO()
    Dim ws As Worksheet
    Set ws = Sheets("MAKRO")

    ' Usuń wszystkie przyciski z arkusza (jeśli mają prefix "btn")
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.Name Like "btnDo_*" Or shp.Name Like "btnReject_*" Then
                shp.Delete
            End If
        End If
    Next shp

    ' Wyczyść zawartość od wiersza 2 w dół
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ws.Rows("2:" & lastRow).ClearContents
    End If

    MsgBox "Zakładka MAKRO została wyczyszczona wraz z przyciskami.", vbInformation
    End Sub
