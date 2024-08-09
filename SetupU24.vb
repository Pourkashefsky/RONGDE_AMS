Private Sub ComEdit_Click()
    List1.List(List1.ListIndex) = Left(LabelNo.Caption, 1) & vbTab & _
        Right(LabelNo.Caption, 2) & vbTab & Val(TextDAU.Text) & vbTab & _
        Val(TextDAL.Text) & vbTab & ComboDispImg.Text
End Sub

Private Sub ComSave_Click()
    Call List2File(List1, App.Path & "\SetU.ini", LabTital.Caption)
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer, N As Integer, S As String
    For N = 0 To 0
        For i = 0 To 23
            S = N & vbTab & Format(i, "00") & vbTab & _
                U24Data(N, i).DAU & vbTab & U24Data(N, i).DAL & vbTab & _
                U24Data(N, i).DispImg
            List1.AddItem S
        Next i
    Next N
    For N = 0 To 9
        For i = 0 To 23
            ComboDispImg.AddItem N & "-" & Format(i, "00")
        Next i
    Next N
    ComboDispImg.AddItem "9-99"
    List1.ListIndex = 0
End Sub

Private Sub List1_Click()
    Call Str2Array(List1.Text)
    If UBound(OutStr) = 4 Then
        LabelNo.Caption = Val(OutStr(0)) & "-" & Format(Val(OutStr(1)), "00")
        TextDAU.Text = OutStr(2): TextDAL.Text = OutStr(3)
        ComboDispImg.Text = OutStr(4)
    End If
End Sub

Private Sub TextPassW_Change()
    If TextPassW.Text = Password Then
        ComSave.Enabled = True
    Else
        ComSave.Enabled = False
    End If
End Sub
