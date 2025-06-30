' 🔍 Znajdź pierwszy task bez przypisanego właściciela
Dim sourceRow As Long: sourceRow = 0
Dim j As Long
For j = 2 To wsSource.Cells(wsSource.Rows.Count, colTask).End(xlUp).Row
    If wsSource.Cells(j, colOwner).Value = "" Then
        sourceRow = j
        Exit For
    End If
Next j

If sourceRow = 0 Then
    MsgBox "Brak wolnych tasków w scope '" & taskName & "'", vbExclamation
    GoTo SkipTask
End If

taskID = wsSource.Cells(sourceRow, colTask).Value
dueDate = wsSource.Cells(sourceRow, colDue).Value


Dim sourceRow As Long: sourceRow = 0
Dim j As Long
For j = 2 To wsSource.Cells(wsSource.Rows.Count, colTaskID).End(xlUp).Row
    If wsSource.Cells(j, colOwner).Value = "" Then
        sourceRow = j
        Exit For
    End If
Next j

If sourceRow = 0 Then
    MsgBox "Brak kolejnych wolnych tasków w scope '" & scopeName & "'", vbInformation
    Exit Sub
End If

newTaskID = wsSource.Cells(sourceRow, colTaskID).Value
newDueDate = wsSource.Cells(sourceRow, colDue).Value
