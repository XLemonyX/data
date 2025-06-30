funkcja

' Szuka numeru kolumny po nazwie nagłówka w zadanym wierszu

Function GetColumnIndex(ws As Worksheet, headerName As String, headerRow As Long) As Long
    Dim col As Long
    For col = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(headerRow, col).Value) = headerName Then
            GetColumnIndex = col
            Exit Function
        End If
    Next col
    GetColumnIndex = 0 ' 0 = nie znaleziono
End Function

funkcja

' Szuka najnowszego wpisu w Assigned_Tasks dla danego Task ID

Function FindLatestAssignmentInfo(taskID As String) As String
    Dim ws As Worksheet: Set ws = Sheets("Assigned_Tasks")
    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value = taskID Then
            FindLatestAssignmentInfo = "to be verified – case already assigned to " & _
                ws.Cells(i, 4).Value & " on " & ws.Cells(i, 5).Value & " from scope " & ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i
    FindLatestAssignmentInfo = ""
End Function


GŁÓWNE MAKRO

' Makro przypisujące taski z zakresu do arkusza MAKRO (wcześniej Example)

Sub AssignTasks(clientName As String, scopeType As String, callingForm As Object)
    Dim wsScope As Worksheet, wsTasks As Worksheet, wsEmp As Worksheet, wsAssigned As Worksheet
    Set wsScope = Sheets("Scope")
    Set wsTasks = Sheets("MAKRO") ' dawny Example
    Set wsEmp = Sheets("Employees")
    Set wsAssigned = Sheets("Assigned_Tasks")

    ' Znajdź osobę wykluczoną
    Dim exclusion As String
    Dim empRow As Range
    Set empRow = wsEmp.Range("A:A").Find(clientName, LookIn:=xlValues, LookAt:=xlWhole)
    If empRow Is Nothing Then MsgBox "Client not found.": Exit Sub
    exclusion = empRow.Offset(0, 1).Value

    ' Wybierz typ Scope
    Dim startCell As Range
    If scopeType = "Standard" Then
        Set startCell = wsScope.Range("C8")
    Else
        Set startCell = wsScope.Range("F8")
    End If

    Dim r As Long: r = 0
    Dim taskName As String, taskCount As Integer

    Do While startCell.Offset(r, 0).Value <> ""
        taskName = startCell.Offset(r, 0).Value
        taskCount = Val(startCell.Offset(r, 1).Value)
        If taskName = exclusion Then
            taskName = "OEDD": taskCount = taskCount + 1
        End If

        Dim wsSource As Worksheet
        On Error Resume Next
        Set wsSource = Sheets(taskName)
        On Error GoTo 0
        If wsSource Is Nothing Then MsgBox "Brak zakładki: " & taskName: Exit Sub

        ' Pobierz indeksy kolumn z zakładki source
        Dim colTask As Long, colDue As Long, colOwner As Long
        colTask = GetColumnIndex(wsSource, "Task", 1)
        colDue = GetColumnIndex(wsSource, "Due Date", 1)
        colOwner = GetColumnIndex(wsSource, "FINAL OWNER", 1)
        If colTask * colDue * colOwner = 0 Then MsgBox "Brakuje kolumn w " & taskName: Exit Sub

        Dim i As Integer
        For i = 1 To taskCount
            ' Znajdź nowy wiersz do wpisu w zakładce MAKRO
            Dim taskID As String, dueDate As Variant
            Dim sourceRow As Long
            sourceRow = wsSource.Cells(wsSource.Rows.Count, colTask).End(xlUp).Row + 1
            taskID = wsSource.Cells(sourceRow, colTask).Value
            dueDate = wsSource.Cells(sourceRow, colDue).Value

            ' Znajdź pierwszy pusty wiersz w MAKRO
            Dim targetRow As Long
            targetRow = wsTasks.Cells(wsTasks.Rows.Count, 1).End(xlUp).Row + 1

            ' Sprawdź czy task był już przypisany
            Dim verifyComment As String
            verifyComment = FindLatestAssignmentInfo(taskID)

            wsTasks.Cells(targetRow, 1).Value = taskID
            wsTasks.Cells(targetRow, 2).Value = dueDate
            wsTasks.Cells(targetRow, 3).Value = "" ' Start Date – do uzupełnienia później
            wsTasks.Cells(targetRow, 4).Value = taskName

            If verifyComment <> "" Then
                wsTasks.Cells(targetRow, 5).Value = verifyComment
                wsTasks.Cells(targetRow, 6).Value = "to verify"
                AddVerificationButtons wsTasks, targetRow
            Else
                wsTasks.Cells(targetRow, 5).Value = clientName
                wsTasks.Cells(targetRow, 6).Value = "Assigned"

                ' Zapisz do historii Assigned_Tasks
                Dim histRow As Long
                histRow = wsAssigned.Cells(wsAssigned.Rows.Count, 1).End(xlUp).Row + 1
                wsAssigned.Cells(histRow, 1).Value = taskID
                wsAssigned.Cells(histRow, 2).Value = dueDate
                wsAssigned.Cells(histRow, 3).Value = taskName
                wsAssigned.Cells(histRow, 4).Value = clientName
                wsAssigned.Cells(histRow, 5).Value = Date
                wsAssigned.Cells(histRow, 6).Value = "Assigned"
            End If
        Next i

        r = r + 1
    Loop

    MsgBox "Taski przypisane do zakładki MAKRO!", vbInformation
    Unload callingForm
End Sub

DODAWANIE PRZYCISKÓW

Sub AddVerificationButtons(ws As Worksheet, rowIndex As Long)

    Dim leftPos As Double, topPos As Double
    topPos = ws.Cells(rowIndex, 6).Top
    leftPos = ws.Cells(rowIndex, 6).Left + ws.Columns(6).Width * 6

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

obsługaa przzycisków

Sub DoTask()
    Dim wsM As Worksheet, wsA As Worksheet
    Set wsM = Sheets("MAKRO")
    Set wsA = Sheets("Assigned_Tasks")

    Dim rowIdx As Long: rowIdx = ActiveCell.Row
    wsM.Cells(rowIdx, 5).Value = Application.UserName
    wsM.Cells(rowIdx, 6).Value = "assigned after manual verification"

    Dim logRow As Long: logRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    With wsA
        .Cells(logRow, 1).Value = wsM.Cells(rowIdx, 1).Value
        .Cells(logRow, 2).Value = wsM.Cells(rowIdx, 2).Value
        .Cells(logRow, 3).Value = wsM.Cells(rowIdx, 4).Value
        .Cells(logRow, 4).Value = Application.UserName
        .Cells(logRow, 5).Value = Date
        .Cells(logRow, 6).Value = "assigned after manual verification"
    End With
End Sub

Sub RejectTask()
    Dim wsM As Worksheet, wsA As Worksheet, wsSource As Worksheet
    Set wsM = Sheets("MAKRO")                  ' Arkusz z zadaniami dystrybuowanymi
    Set wsA = Sheets("Assigned_Tasks")         ' Historia przypisań

    Dim rowIdx As Long: rowIdx = ActiveCell.Row

    Dim taskID As String: taskID = wsM.Cells(rowIdx, 1).Value
    Dim dueDate As Variant: dueDate = wsM.Cells(rowIdx, 2).Value
    Dim scopeName As String: scopeName = wsM.Cells(rowIdx, 4).Value

    ' 📝 Zapytaj o powód odrzucenia
    Dim reason As String
    reason = InputBox("Podaj powód odrzucenia:", "Reject Task")
    If reason = "" Then Exit Sub

    ' ✍️ Dodaj wpis do Assigned_Tasks jako „Rejected”
    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    With wsA
        .Cells(lastRow, 1).Value = taskID
        .Cells(lastRow, 2).Value = dueDate
        .Cells(lastRow, 3).Value = scopeName
        .Cells(lastRow, 4).Value = "case rejected by distributor due to: " & reason
        .Cells(lastRow, 5).Value = "" ' brak daty
        .Cells(lastRow, 6).Value = "Rejected"
    End With

    ' 🧹 Usuń wiersz z MAKRO
    wsM.Rows(rowIdx).Delete

    ' 🔁 Załaduj nowy task z tego samego scope, jeśli istnieje
    On Error Resume Next
    Set wsSource = Sheets(scopeName)
    On Error GoTo 0
    If wsSource Is Nothing Then Exit Sub

    Dim colTaskID As Long, colDue As Long
    colTaskID = GetColumnIndex(wsSource, "Task", 1)
    colDue = GetColumnIndex(wsSource, "Due Date", 1)
    If colTaskID * colDue = 0 Then Exit Sub

    ' Znajdź nowy task do załadowania
    Dim sourceRow As Long
    sourceRow = wsSource.Cells(wsSource.Rows.Count, colTaskID).End(xlUp).Row + 1
    Dim newTaskID As String, newDueDate As Variant
    newTaskID = wsSource.Cells(sourceRow, colTaskID).Value
    newDueDate = wsSource.Cells(sourceRow, colDue).Value

    ' Sprawdź, czy to duplikat
    Dim comment As String
    comment = FindLatestAssignmentInfo(newTaskID)

    Dim newRow As Long
    newRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row + 1

    wsM.Cells(newRow, 1).Value = newTaskID
    wsM.Cells(newRow, 2).Value = newDueDate
    wsM.Cells(newRow, 3).Value = ""
    wsM.Cells(newRow, 4).Value = scopeName

    If comment <> "" Then
        wsM.Cells(newRow, 5).Value = comment
        wsM.Cells(newRow, 6).Value = "to verify"
        AddVerificationButtons wsM, newRow
    Else
        wsM.Cells(newRow, 5).Value = "" ' dystrybutor może wpisać ręcznie
        wsM.Cells(newRow, 6).Value = "Assigned"
    End If
End Sub

