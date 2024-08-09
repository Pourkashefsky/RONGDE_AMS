Private Sub ComAll_Click()
    TextFind.Text = "Select * From [LEVEL]"
    ComSearch_Click
End Sub

Private Sub ComboA_Click()
Dim i As Integer, j As Integer
    Select Case ComboA.Text     '更新查询条件选择框
    Case "Time"
        ComboC.Clear
        ComboC.AddItem "Today"
        ComboC.AddItem "The Month"
        ComboC.AddItem "The Year"
    Case "Name"
        ComboC.Clear
        For i = 0 To 9
            For j = 0 To 31
                ComboC.AddItem BinData(i, j).Name
            Next j
        Next i
        For i = 0 To 9
            For j = 0 To 23
                ComboC.AddItem MoniData(i, j).Name
            Next j
        Next i
    End Select
    ComboC.ListIndex = 0
    
    MakeFindString              '组建检索语句
End Sub

Private Sub ComboB_Click()
    MakeFindString              '组建检索语句
End Sub

Private Sub ComboC_Click()
    MakeFindString              '组建检索语句
End Sub

Private Sub ComboD_Click()
    MakeFindString              '组建检索语句
End Sub

Private Sub ComboE_Click()
    MakeFindString              '组建检索语句
End Sub

Private Sub ComClear_Click()        '清空数据
    'Adodc1.RecordSource = "Delete * From [LEVEL]"
    'DoEvents
    'Call Sleep(1)
    
    'Adodc1.RecordSource = "Select * From [LEVEL]"
    'Adodc1.Refresh
    'DataGrid1.Refresh
    'DoEvents
    Call ADODel(Adodc1, "[LEVEL]")
    DoEvents
    Call Sleep(1)
    Adodc1.Refresh
    DataGrid1.Refresh
    DoEvents
End Sub

Private Sub ComClose_Click()
    Unload Me
End Sub

Private Sub ComPrint_Click()        '打印报警记录
ReDim PrintCol(0 To 2) As Integer   '打印列号数组
Dim i As Integer, Ss As String
    For i = 0 To 2                  '打印第0,1,2列
        PrintCol(i) = i
    Next
    Ss = "Alarm List"               '表头标题
    Pp = 0                          '起始页码
    Call ADOPrint(Adodc1, PrintCol(), Ss, 0, Adodc1.Recordset.RecordCount - 1, 1, 55, 2500)
End Sub

Private Sub ComSearch_Click()       '数据查询
On Error GoTo ErrMsg
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = TextFind.Text
    Adodc1.Refresh
    DataGrid1.Columns.Item(0).Width = 2000
    DataGrid1.Columns.Item(1).Width = 4000
    DataGrid1.Columns.Item(2).Width = 1600
    Exit Sub
ErrMsg:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\ALARM.mdb"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "Select * From [LEVEL]"
    Adodc1.Refresh
    
    ComboA.ListIndex = 0
    ComboB.ListIndex = 0
    ComboD.ListIndex = 0
    ComboE.ListIndex = 0
    DataGrid1.Columns.Item(0).Width = 2000
    DataGrid1.Columns.Item(1).Width = 4000
    DataGrid1.Columns.Item(2).Width = 1600
End Sub

Private Sub TextPass_Change()
    If TextPass.Text = "rongded" Or TextPass.Text = "RONGDED" Then
        ComClear.Enabled = True
    Else
        ComClear.Enabled = False
    End If
End Sub

Sub MakeFindString()        '组建检索语句
Dim txtFind As String, txtOrder As String
Dim TimeS As String, TimeE As String
    txtFind = "[" & ComboA.Text & "] " & ComboB.Text
    If ComboE.ListIndex = 0 Then txtOrder = " Order By [" & ComboD.Text & "] ASC"
    If ComboE.ListIndex = 1 Then txtOrder = " Order By [" & ComboD.Text & "] DESC"
    
    Select Case ComboA.Text
        Case "Time"
            Select Case ComboC.Text
            Case "Today"        '当日记录
                txtFind = "Select * From [LEVEL] Where [Time] Like '" & Format(Now, tFms1) & "%'" & txtOrder
            Case "The Month"    '当月记录
                txtFind = "Select * From [LEVEL] Where [Time] Like '%" & Format(Now, tFms2) & "%'" & txtOrder
            Case "The Year"     '当年记录
                txtFind = "Select * From [LEVEL] Where [Time] Like '%" & Format(Now, "yyyy") & "%'" & txtOrder
            End Select
        Case "Name"
            txtFind = "Select * From [LEVEL] Where " & txtFind & " '" & ComboC.Text & "%'" & txtOrder
    End Select
    
    TextFind.Text = txtFind
End Sub

