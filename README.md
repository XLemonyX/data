    Private Sub UserForm_Initialize()
    'Wypenij boxa lista z zakladki employees
    Dim lastRow As Long
    lastRow = Sheets("Employees").Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        ComboBox1.AddItem Sheets("Employees").Cells(i, 1).Value
    Next i
    
    'Scope options
    ComboBox2.AddItem "Standard"
    ComboBox2.AddItem "Overtime"
    
    End Sub

    Private Sub CommandButton1_Click()
    Call AssignTasks(ComboBox1.Value, ComboBox2.Value, Me)
    End Sub
