' ======= FUNKCJA: Znajdź kolumnę po nazwie nagłówka =======

    Function GetColumnIndex(ws As Worksheet, headerName As String, headerRow As Long) As Long
    Dim col As Long
    For col = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(headerRow, col).Value) = headerName Then
            GetColumnIndex = col
            Exit Function
        End If
    Next col
    GetColumnIndex = 0
End Function

' ======= FUNKCJA: Znajdź ostatnie przypisanie danego taska =======


    Function FindLatestAssignmentInfo(taskID As String) As String
    Dim ws As Worksheet: Set ws = Sheets("Assigned_Tasks")
    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value = taskID Then
            FindLatestAssignmentInfo = ws.Cells(i, 4).Value & " on " & ws.Cells(i, 5).Value & " from scope " & ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i
    FindLatestAssignmentInfo = ""
End Function

' ======= PROCEDURA: Dodaj przyciski Do Task / Reject =======

    Sub AddVerificationButtons(ws As Worksheet, rowIndex As Long)
    Dim leftPos As Double, topPos As Double
    topPos = ws.Cells(rowIndex, 6).Top
    leftPos = ws.Cells(rowIndex, 6).Left + 70

    With ws.Buttons.Add(leftPos, topPos, 65, 18)
        .Caption = "✅ Do Task"
        .OnAction = "DoTask"
        .Name = "btnDo_" & rowIndex
    End With

    With ws.Buttons.Add(leftPos + 70, topPos, 65, 18)
        .Caption = "❌ Reject"
        .OnAction = "RejectTask"
        .Name = "btnReject_" & rowIndex
    End With
End Sub

' ======= MAKRO: Przypisanie tasków do MAKRO + logowanie =======

    Sub AssignTasks(clientName As String, scopeName As String, callingForm As Object)
    Dim wsSource As Worksheet, wsM As Worksheet, wsA As Worksheet
    Set wsSource = Sheets(scopeName)
    Set wsM = Sheets("MAKRO")
    Set wsA = Sheets("Assigned_Tasks")

    Dim colTask As Long, colDue As Long, colOwner As Long
    colTask = GetColumnIndex(wsSource, "Task", 1)
    colDue = GetColumnIndex(wsSource, "Due Date", 1)
    colOwner = GetColumnIndex(wsSource, "FINAL OWNER", 1)
    If colTask * colDue * colOwner = 0 Then
        MsgBox "Brakuje kolumn w arkuszu " & scopeName, vbExclamation
        Exit Sub
    End If

    Dim countToAssign As Long: countToAssign = 3 ' lub dynamicznie
    Dim i As Long, foundRow As Long, assigned As Long: assigned = 0

    For i = 2 To wsSource.Cells(wsSource.Rows.Count, colTask).End(xlUp).Row
        If wsSource.Cells(i, colOwner).Value = "" Then
            Dim taskID As String: taskID = wsSource.Cells(i, colTask).Value
            Dim dueDate As Variant: dueDate = wsSource.Cells(i, colDue).Value
            Dim comment As String: comment = FindLatestAssignmentInfo(taskID)

            ' Dodaj do historii
            Dim logRow As Long: logRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
            With wsA
                .Cells(logRow, 1).Value = taskID
                .Cells(logRow, 2).Value = dueDate
                .Cells(logRow, 3).Value = scopeName
                If comment = "" Then
                    .Cells(logRow, 4).Value = clientName
                    .Cells(logRow, 5).Value = Date
                    .Cells(logRow, 6).Value = "Assigned"
                Else
                    .Cells(logRow, 4).Value = "to be verified – case already assigned to: " & comment
                    .Cells(logRow, 5).Value = ""
                    .Cells(logRow, 6).Value = "Duplicate – verify manually"
                End If
            End With

            ' Dodaj do MAKRO
            Dim mRow As Long: mRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row + 1
            wsM.Cells(mRow, 1).Value = taskID
            wsM.Cells(mRow, 2).Value = dueDate
            wsM.Cells(mRow, 3).Value = ""
            wsM.Cells(mRow, 4).Value = scopeName
            If comment = "" Then
                wsM.Cells(mRow, 5).Value = clientName
                wsM.Cells(mRow, 6).Value = "Assigned"
            Else
                wsM.Cells(mRow, 5).Value = "to be verified – case already assigned to: " & comment
                wsM.Cells(mRow, 6).Value = "to verify"
                AddVerificationButtons wsM, mRow
            End If

            assigned = assigned + 1
            If assigned = countToAssign Then Exit For
        End If
    Next i

    If assigned = 0 Then MsgBox "Brak wolnych tasków w scope '" & scopeName & "'.", vbInformation
    Unload callingForm
End Sub

' ======= OBSŁUGA PRZYCISKU: Do Task =======

    Sub DoTask()
    Dim wsM As Worksheet, wsA As Worksheet
    Set wsM = Sheets("MAKRO")
    Set wsA = Sheets("Assigned_Tasks")

    Dim rowIdx As Long: rowIdx = ActiveCell.Row
    Dim taskID As String: taskID = wsM.Cells(rowIdx, 1).Value
    Dim dueDate As Variant: dueDate = wsM.Cells(rowIdx, 2).Value
    Dim scopeName As String: scopeName = wsM.Cells(rowIdx, 4).Value

    ' Zapisz do Assigned_Tasks
    Dim logRow As Long: logRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    With wsA
        .Cells(logRow, 1).Value = taskID
        .Cells(logRow, 2).Value = dueDate
        .Cells(logRow, 3).Value = scopeName
        .Cells(logRow, 4).Value = Application.UserName
        .Cells(logRow, 5).Value = Date
        .Cells(logRow, 6).Value = "assigned after manual verification"
    End With

    ' Zmień dane w MAKRO (opcjonalnie)
    wsM.Cells(rowIdx, 5).Value = Application.UserName
    wsM.Cells(rowIdx, 6).Value = "✅ assigned manually"
End Sub

' ======= OBSŁUGA PRZYCISKU: Reject =======

    Sub RejectTask()
    Dim wsM As Worksheet, wsA As Worksheet, wsSource As Worksheet
    Set wsM = Sheets("MAKRO")
    Set wsA = Sheets("Assigned_Tasks")

    Dim rowIdx As Long: rowIdx = ActiveCell.Row
    Dim taskID As String: taskID = wsM.Cells(rowIdx, 1).Value
    Dim dueDate As Variant: dueDate = wsM.Cells(rowIdx, 2).Value
    Dim scopeName As String: scopeName = wsM.Cells(rowIdx, 4).Value

    Dim reason As String
    reason = InputBox("Podaj powód odrzucenia:", "Reject Task")
    If reason = "" Then Exit Sub

    Dim logRow As Long: logRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    With wsA
        .Cells(logRow, 1).Value = taskID
        .Cells(logRow, 2).Value = dueDate
        .Cells(logRow, 3).Value = scopeName
        .Cells(logRow, 4).Value = "case rejected by distributor due to: " & reason
        .Cells(logRow, 5).Value = "" ' brak daty
        .Cells(logRow, 6).Value = "Rejected"
    End With

    wsM.Rows(rowIdx).Delete

       ' Dodaj nowy task z tego samego scope
    On Error Resume Next
    Set wsSource = Sheets(scopeName)
    On Error GoTo 0
    If wsSource Is Nothing Then Exit Sub

    Dim colTaskID As Long, colDue As Long, colOwner As Long
    colTaskID = GetColumnIndex(wsSource, "Task", 1)
    colDue = GetColumnIndex(wsSource, "Due Date", 1)
    colOwner = GetColumnIndex(wsSource, "FINAL OWNER", 1)
    If colTaskID * colDue * colOwner = 0 Then Exit Sub

    ' 🔍 Znajdź pierwszy wolny task (czyli taki, gdzie FINAL OWNER = "")
    Dim j As Long, sourceRow As Long: sourceRow = 0
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

    ' Pobierz dane nowego taska
    Dim newTaskID As String: newTaskID = wsSource.Cells(sourceRow, colTaskID).Value
    Dim newDueDate As Variant: newDueDate = wsSource.Cells(sourceRow, colDue).Value
    Dim newComment As String: newComment = FindLatestAssignmentInfo(newTaskID)

    ' Zapisz nowy wpis do Assigned_Tasks
    Dim newLogRow As Long: newLogRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    With wsA
        .Cells(newLogRow, 1).Value = newTaskID
        .Cells(newLogRow, 2).Value = newDueDate
        .Cells(newLogRow, 3).Value = scopeName
        If newComment = "" Then
            .Cells(newLogRow, 4).Value = ""
            .Cells(newLogRow, 5).Value = ""
            .Cells(newLogRow, 6).Value = "Assigned"
        Else
            .Cells(newLogRow, 4).Value = "to be verified – case already assigned to: " & newComment
            .Cells(newLogRow, 5).Value = ""
            .Cells(newLogRow, 6).Value = "Duplicate – verify manually"
        End If
    End With

    ' Dodaj nowy task do MAKRO
    Dim mRow As Long: mRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row + 1
    wsM.Cells(mRow, 1).Value = newTaskID
    wsM.Cells(mRow, 2).Value = newDueDate
    wsM.Cells(mRow, 3).Value = ""
    wsM.Cells(mRow, 4).Value = scopeName
    If newComment <> "" Then
        wsM.Cells(mRow, 5).Value = "to be verified – case already assigned to: " & newComment
        wsM.Cells(mRow, 6).Value = "to verify"
        AddVerificationButtons wsM, mRow
    Else
        wsM.Cells(mRow, 5).Value = ""
        wsM.Cells(mRow, 6).Value = "Assigned"
    End If
End Sub

