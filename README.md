    Sub AssignTasks(clientName As String, scopeType As String, callingForm As Object)
    Dim wsScope As Worksheet, wsMAKRO As Worksheet, wsTasks As Worksheet, wsLog As Worksheet
    Set wsScope = Sheets("Scope")
    Set wsMAKRO = Sheets("MAKRO")
    Set wsLog = Sheets("Assigned_Tasks")
    
    ' Sprawdź wykluczenia
    Dim wsEmp As Worksheet: Set wsEmp = Sheets("Employees")
    Dim exclusion As String
    Dim empRow As Range: Set empRow = wsEmp.Range("A:A").Find(clientName, LookIn:=xlValues)
    If Not empRow Is Nothing Then exclusion = empRow.Offset(0, 1).Value
    
    ' Określ start kolumny zakresu (Standard vs Overtime)
    Dim scopeStart As Range
    If scopeType = "Standard" Then
        Set scopeStart = wsScope.Range("C8")
    Else
        Set scopeStart = wsScope.Range("F8")
    End If

    Dim taskName As String, taskCount As Integer, r As Long
    r = 0

    Do While scopeStart.Offset(r, 0).Value <> ""
        taskName = scopeStart.Offset(r, 0).Value
        taskCount = scopeStart.Offset(r, 1).Value

        ' Obsłuż wykluczenie
        If taskName = exclusion Then
            taskName = "OEDD"
            taskCount = taskCount + 1
        End If
        
        On Error Resume Next
        Set wsTasks = Sheets(taskName)
        On Error GoTo 0
        If wsTasks Is Nothing Then
            MsgBox "Brak zakładki: " & taskName, vbExclamation
            r = r + 1
            Continue Do
        End If

        Dim colTask As Long, colDue As Long, colOwner As Long
        colTask = GetColumnIndex(wsTasks, "Task", 1)
        colDue = GetColumnIndex(wsTasks, "Due Date", 1)
        colOwner = GetColumnIndex(wsTasks, "FINAL OWNER", 1)
        If colTask * colDue * colOwner = 0 Then
            MsgBox "Brakuje kolumn w zakładce " & taskName, vbExclamation
            r = r + 1
            Continue Do
        End If

        Dim i As Long, j As Long, foundRow As Long, assigned As Long: assigned = 0

        For j = 2 To wsTasks.Cells(wsTasks.Rows.Count, colTask).End(xlUp).Row
            If wsTasks.Cells(j, colOwner).Value = "" Then
                Dim taskID As String: taskID = wsTasks.Cells(j, colTask).Value
                Dim dueDate As Variant: dueDate = wsTasks.Cells(j, colDue).Value
                Dim comment As String: comment = FindLatestAssignmentInfo(taskID)
                
                ' Zapisz do Assigned_Tasks
                Dim logRow As Long: logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
                With wsLog
                    .Cells(logRow, 1).Value = taskID
                    .Cells(logRow, 2).Value = dueDate
                    .Cells(logRow, 3).Value = taskName
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
                
                ' Dodaj do arkusza MAKRO
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
                
                assigned = assigned + 1
                If assigned = taskCount Then Exit For
            End If
        Next j

        r = r + 1
    Loop

    MsgBox "Przypisano zadania dla: " & clientName, vbInformation
    Unload callingForm
    End Sub
