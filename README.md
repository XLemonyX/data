Private Sub AddClient_Click()
 Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("clients")

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(1)

    Dim existingIDs As Range
    Set existingIDs = tbl.ListColumns("CustomerID").DataBodyRange

    Dim newID As String: newID = Trim(Me.TextBox1.Value)
    Dim newName As String: newName = Trim(Me.TextBox2.Value)
    Dim newCountry As String: newCountry = Trim(Me.TextBox3.Value)

    ' Check if ID is empty
    If newID = "" Or newName = "" Or newCountry = "" Then
        MsgBox "Please fill in all fields.", vbExclamation
        Exit Sub
    End If

    ' Check if ID already exists
    If Application.WorksheetFunction.CountIf(existingIDs, newID) > 0 Then
        MsgBox "Client already exists!", vbCritical
        Exit Sub
    End If

    ' Add new row
    With tbl.ListRows.Add
        .Range(1, 1).Value = newID
        .Range(1, 2).Value = newName
        .Range(1, 3).Value = newCountry
    End With

    MsgBox "New client added successfully!", vbInformation
    Unload Me
End Sub
