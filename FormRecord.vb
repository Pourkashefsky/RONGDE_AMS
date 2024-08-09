Private Sub ComAll_Click()
    TextFind.Text = "Select * From [LEVEL]"
    ComSearch_Click
End Sub

Private Sub ComboA_Click()
Dim i As Integer, j As Integer
    Select Case ComboA.Text     '���²�ѯ����ѡ���
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
    
    MakeFindString              '�齨�������
End Sub

Private Sub ComboB_Click()
    MakeFindString              '�齨�������
End Sub

Private Sub ComboC_Click()
    MakeFindString              '�齨�������
End Sub

Private Sub ComboD_Click()
    MakeFindString              '�齨�������
End Sub

Private Sub ComboE_Click()
    MakeFindString              '�齨�������
End Sub

Private Sub ComClear_Click()        '�������
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

Private Sub ComPrint_Click()        '��ӡ������¼
ReDim PrintCol(0 To 2) As Integer   '��ӡ�к�����
Dim i As Integer, Ss As String
    For i = 0 To 2                  '��ӡ��0,1,2��
        PrintCol(i) = i
    Next
    Ss = "Alarm List"               '��ͷ����
    Pp = 0                          '��ʼҳ��
    Call ADOPrint(Adodc1, PrintCol(), Ss, 0, Adodc1.Recordset.RecordCount - 1, 1, 55, 2500)
End Sub

Private Sub ComSearch_Click()       '���ݲ�ѯ
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

Sub MakeFindString()        '�齨�������
Dim txtFind As String, txtOrder As String
Dim TimeS As String, TimeE As String
    txtFind = "[" & ComboA.Text & "] " & ComboB.Text
    If ComboE.ListIndex = 0 Then txtOrder = " Order By [" & ComboD.Text & "] ASC"
    If ComboE.ListIndex = 1 Then txtOrder = " Order By [" & ComboD.Text & "] DESC"
    
    Select Case ComboA.Text
        Case "Time"
            Select Case ComboC.Text
            Case "Today"        '���ռ�¼
                txtFind = "Select * From [LEVEL] Where [Time] Like '" & Format(Now, tFms1) & "%'" & txtOrder
            Case "The Month"    '���¼�¼
                txtFind = "Select * From [LEVEL] Where [Time] Like '%" & Format(Now, tFms2) & "%'" & txtOrder
            Case "The Year"     '�����¼
                txtFind = "Select * From [LEVEL] Where [Time] Like '%" & Format(Now, "yyyy") & "%'" & txtOrder
            End Select
        Case "Name"
            txtFind = "Select * From [LEVEL] Where " & txtFind & " '" & ComboC.Text & "%'" & txtOrder
    End Select
    
    TextFind.Text = txtFind
End Sub

