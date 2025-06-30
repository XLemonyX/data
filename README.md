    Sub ClearMAKRO()
    Dim ws As Worksheet
    Set ws = Sheets("MAKRO")

    ' 1️⃣ Usuń WSZYSTKIE Shape'y (w tym przyciski Form Control)
    Dim shp As Shape
    Dim shapeIndex As Long

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(shapeIndex).Delete
        On Error GoTo 0
    Next shapeIndex

    ' 2️⃣ Usuń WSZYSTKIE obiekty ActiveX (jeśli jakieś istnieją)
    Dim obj As OLEObject
    For Each obj In ws.OLEObjects
        On Error Resume Next
        obj.Delete
        On Error GoTo 0
    Next obj

    ' 3️⃣ Wyczyść dane od wiersza 2 w dół
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ws.Rows("2:" & lastRow).ClearContents
    End If

    MsgBox "✅ Zakładka 'MAKRO' całkowicie wyczyszczona – dane i przyciski zniknęły!", vbInformation
    End Sub
