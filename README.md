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

    Function FindLatestAssignmentInfo(taskID As String) As String
    Dim wsLog As Worksheet: Set wsLog = Sheets("Assigned_Tasks")
    Dim lastRow As Long: lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    Dim i As Long

    For i = lastRow To 2 Step -1
        If wsLog.Cells(i, 1).Value = taskID Then
            Dim assignee As String: assignee = wsLog.Cells(i, 4).Value
            Dim assignDate As Variant: assignDate = wsLog.Cells(i, 5).Value
            Dim scopeName As String: scopeName = wsLog.Cells(i, 3).Value
            If assignee <> "" And assignDate <> "" Then
                FindLatestAssignmentInfo = assignee & " on " & Format(assignDate, "yyyy-mm-dd") & " from scope " & scopeName
            End If
            Exit Function
        End If
    Next i

    FindLatestAssignmentInfo = ""
    End Function

    Sub AddVerificationButtons(ws As Worksheet, rowIndex As Long)
    Dim btnDo As Button, btnReject As Button
    Dim topPos As Double: topPos = ws.Rows(rowIndex).Top
    Dim leftDo As Double: leftDo = ws.Columns(7).Left
    Dim leftReject As Double: leftReject = ws.Columns(8).Left
    Dim btnHeight As Double: btnHeight = 18
    Dim btnWidth As Double: btnWidth = 60

    Set btnDo = ws.Buttons.Add(leftDo, topPos, btnWidth, btnHeight)
    With btnDo
        .OnAction = "DoTask"
        .Caption = "✅ Do Task"
        .Name = "btnDo_" & rowIndex
    End With

    Set btnReject = ws.Buttons.Add(leftReject, topPos, btnWidth, btnHeight)
    With btnReject
        .OnAction = "RejectTask"
        .Caption = "❌ Reject"
        .Name = "btnReject_" & rowIndex
    End With
    End Sub
