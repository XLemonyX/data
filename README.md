Sub AssignTasks(clientName As String, scopeType As String, callingForm As Object)
    Dim wsScope As Worksheet, wsExample As Worksheet, wsEmp As Worksheet
    Set wsScope = Sheets("Scope")
    Set wsExample = Sheets("Example")
    Set wsEmp = Sheets("Employees")
    
    'sprawdzenie wykluczen
    Dim exclusion As String
    Dim empRow As Range
    Set empRow = wsEmp.Range("A:A").Find(clientName, LookIn:=xlValues)
    If Not empRow Is Nothing Then
        exclusion = empRow.Offset(0, 1).Value
    End If
    
    'okreslenie zakresu tabeli scope
    Dim scopeStart As Range
    If scopeType = "Standard" Then
        Set scopeStart = wsScope.Range("C8")
    Else
        Set scopeStart = wsScope.Range("F8")
    End If
    
    'Przypisywanie taskow
    Dim taskName As String, taskCount As Integer
    Dim r As Long: r = 0
    Do While scopeStart.Offset(r, 0).Value <> ""
        taskName = scopeStart.Offset(r, 0).Value
        taskCount = scopeStart.Offset(r, 1).Value
        
        If taskName = exclusion Then
            'Pomijamy task wykluczony i zamieniamy go na 1xOEDD
            taskName = "OEDD"
            taskCount = taskCount + 1
        End If
        
    'Przypisanie taskow
    
        Dim wsTask As Worksheet
        Set wsTask = Sheets(taskName)
        Dim i As Long, rowInsert As Long
    
        For i = 1 To taskCount
            rowInsert = wsTask.Cells(wsTask.Rows.Count, 6).End(xlUp).Row + 1 'kolumna 6
            wsTask.Cells(rowInsert, 6).Value = clientName
        
        'Wrzucenie podsumowania do Example
        
            Dim summaryRow As Long
            summaryRow = wsExample.Cells(wsExample.Rows.Count, 1).End(xlUp).Row + 1
            wsExample.Cells(summaryRow, 1).Value = taskName
            wsExample.Cells(summaryRow, 2).Value = wsTask.Cells(rowInsert, 1).Value 'Task #
            wsExample.Cells(summaryRow, 3).Value = wsTask.Cells(rowInsert, 2).Value 'Due Date
            wsExample.Cells(summaryRow, 4).Value = clientName
        Next i
    
        r = r + 1
    Loop

    MsgBox "Tasks successfully assigned to " & clientName, vbInformation
    
    Unload callingForm
    
End Sub
