Sub AssignTasks(clientName As String, scopeType As String, callingForm As Object)
    Dim wsScope As Worksheet, wsMAKRO As Worksheet, wsLog As Worksheet, wsEmp As Worksheet
    Set wsScope = Sheets("Scope")
    Set wsMAKRO = Sheets("MAKRO")
    Set wsLog = Sheets("Assigned_Tasks")
    Set wsEmp = Sheets("Employees")

    ' Wykluczenie
    Dim exclusion As String
    Dim empRow As Range: Set empRow = wsEmp.Range("A:A").Find(clientName, LookIn:=xlValues)
    If Not empRow Is Nothing Then exclusion = empRow.Offset(0, 1).Value

    ' Start zakresu scope
    Dim scopeStart As Range
    If scopeType = "Standard" Then
        Set scopeStart = wsScope.Range("C8")
    Else
        Set scopeStart = wsScope.Range("F8")
    End If

    Dim taskName As String, taskCount As Integer, r As Long: r = 0

    Do While scopeStart.Offset(r, 0).Value <> ""
        taskName = scopeStart.Offset(r, 0).Value
        taskCount = scopeStart.Offset(r, 1).Value

        If taskName = exclusion Then
            taskName = "OEDD"
            taskCount = taskCount + 1
        End If

        On Error Resume Next
        Dim wsTasks As Worksheet: Set wsTasks = Sheets(taskName)
        On Error GoTo 0
        If wsTasks Is Nothing Then
            r = r + 1
            GoTo NextScope
        End If

        Dim colTask As Long, colDue As Long, colOwner As Long
        colTask = GetColumnIndex(wsTasks, "Task", 1)
        colDue = GetColumnIndex(wsTasks, "Due Date", 1)
        colOwner = GetColumnIndex(wsTasks, "FINAL OWNER", 1)
        If colTask * colDue * colOwner = 0 Then
            r = r + 1
            GoTo NextScope
        End If

        Dim j As Long, found As Long: found = 0
        For j = 2 To wsTasks.Cells(wsTasks.Rows.Count, colTask).End(xlUp).Row
            If wsTasks.Cells(j, colOwner).Value = "" Then
                Dim taskID As String: taskID = wsTasks.Cells(j, colTask).Value
                Dim dueDate As Variant: dueDate = wsTasks.Cells(j, colDue).Value
                Dim comment As String: comment = FindLatestAssignmentInfo(taskID)

                ' Dodaj do logu
                Dim logRow As Long: logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
                With wsLog
                    .Cells(logRow, 1).Value = taskID
                    .Cells(logRow, 2).Value = dueDate
                    .Cells(logRow, 3).Value = taskName
                    If comment = "" Then
                        .Cells(logRow, 4).Value = clientName
                        .Cells(logRow, 5).Value = Date
                        .Cells(logRow, 6).Value = "Assigned"
                        wsTasks.Cells(j, colOwner).Value = clientName ' TU WRACAMY DO WPISYWANIA
                    Else
                        .Cells(logRow, 4).Value = "to be verified – case already assigned to: " & comment
                        .Cells(logRow, 5).Value = ""
                        .Cells(logRow, 6).Value = "Duplicate – verify manually"
                    End If
                End With

                ' Dodaj do MAKRO
                Dim mRow As Long: mRow = wsMAKRO.Cells(wsMAKRO.Rows.Count, 1).End(xlUp).Row + 1
                wsMAKRO.Cells(mRow, 1).Value = taskID
                wsMAKRO.Cells(mRow, 2).Value = dueDate
                wsMAKRO.Cells(mRow, 3).Value = ""
                wsMAKRO.Cells(mRow, 4).Value = taskName
                If comment = "" Then
                    wsMAKRO.Cells(mRow, 5).Value = clientName
                    wsMAKRO.Cells(mRow, 6).Value = "Assigned"
                Else
                    wsMAKRO.Cells(mRow, 5).Value = "to be verified – case already assigned to: " & comment
                    wsMAKRO.Cells(mRow, 6).Value = "to verify"
                    AddVerificationButtons wsMAKRO, mRow
                End If

                found = found + 1
                If found = taskCount Then Exit For
            End If
        Next j

        r = r + 1
NextScope:
    Loop

    MsgBox "Przypisano zadania dla " & clientName, vbInformation
    Unload callingForm
End Sub
