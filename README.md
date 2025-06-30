    Sub ClearMAKRO()
    Dim ws As Worksheet
    Set ws = Sheets("MAKRO")

    ' 1️⃣ Usuń tylko przyciski zaczynające się od "btn"
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.Name Like "btn*" Then
                On Error Resume Next
                shp.Delete
                On Error GoTo 0
            End If
        End If
    Next shp

    ' 2️⃣ Wyczyść dane od wiersza 2 w dół
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ws.Rows("2:" & lastRow).ClearContents
    End If

    MsgBox "Zakładka MAKRO wyczyszczona 🧼", vbInformation
    End Sub
