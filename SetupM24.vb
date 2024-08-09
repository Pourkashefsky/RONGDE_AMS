Private Sub ComEdit_Click()
    List1.List(List1.ListIndex) = Val(Left(LabelNo.Caption, 1)) & vbTab & _
        Format(Val(Right(LabelNo.Caption, 2)), "00") & vbTab & _
        Format(ComboDelay.ListIndex, "00") & vbTab & ComboG.ListIndex & vbTab & _
        Val(TextADU.Text) & vbTab & Val(TextADC.Text) & vbTab & Val(TextADL.Text) & vbTab & _
        Format(Val(TextDU.Text), "0.00") & vbTab & _
        Format(Val(TextDL.Text), "0.00") & vbTab & _
        TextUnit.Text & vbTab & TextFmtS.Text & vbTab & _
        Format(Val(TextAlmU.Text), "0.00") & vbTab & _
        Format(Val(TextAlmL.Text), "0.00") & vbTab & _
        Val(TextSFU.Text) & vbTab & _
        Val(TextSFL.Text) & vbTab & _
        CheckCt.Value & vbTab & ComboCutImg.Text & vbTab & _
        TextName.Text
End Sub

Private Sub ComMax_Click()
    TextADU.Text = TextAD.Text
End Sub

Private Sub ComCnt_Click()
    TextADC.Text = TextAD.Text
End Sub

Private Sub ComMin_Click()
    TextADL.Text = TextAD.Text
End Sub

Private Sub ComSave_Click()
    Call List2File(List1, App.Path & "\SetM.ini", LabTital.Caption)
    Call LoadSet
    Unload Me
End Sub

Private Sub ComSFL_Click()
    TextSFL.Text = TextAD.Text
End Sub

Private Sub ComSFU_Click()
    TextSFU.Text = TextAD.Text
End Sub

Private Sub Form_Load()
Dim N As Integer, i As Integer, s As String, Ct As Integer
    For i = 0 To Maxdelay
        ComboDelay.AddItem i 'Format(i * 3.38, "00.0 S")
    Next i
    For i = 0 To 8
        ComboG.AddItem i
    Next i
    
    For N = 0 To 9
        For i = 0 To 23
            If MoniData(N, i).Cutout = True Then
                Ct = 1
            Else
                Ct = 0
            End If
            s = N & vbTab & Format(i, "00") & vbTab & _
                Format(MoniData(N, i).Delay, "00") & vbTab & MoniData(N, i).Group & vbTab & _
                MoniData(N, i).ADU & vbTab & MoniData(N, i).ADC & vbTab & MoniData(N, i).ADL & vbTab & _
                Format(MoniData(N, i).DispU, "0.00") & vbTab & _
                Format(MoniData(N, i).DispL, "0.00") & vbTab & _
                MoniData(N, i).Unit & vbTab & MoniData(N, i).FmtS & vbTab & _
                Format(MoniData(N, i).AlmU, "0.00") & vbTab & _
                Format(MoniData(N, i).AlmL, "0.00") & vbTab & _
                MoniData(N, i).SFU & vbTab & MoniData(N, i).SFL & vbTab & _
                Ct & vbTab & MoniData(N, i).CutImg & vbTab & _
                MoniData(N, i).Name
            List1.AddItem s
        Next i
    Next N
    
    For N = 0 To 9
        For i = 0 To 31
            ComboCutImg.AddItem N & "-" & Format(i, "00")
        Next i
    Next N
    ComboCutImg.AddItem "9-99"
    
    List1.ListIndex = 0
End Sub

Private Sub List1_Click()
    Call Str2Array(List1.Text)
    If UBound(OutStr) = 16 Then
        LabelNo.Caption = Val(OutStr(0)) & "-" & Format(Val(OutStr(1)), "00")
        ComboDelay.ListIndex = Val(OutStr(2)): ComboG.ListIndex = Val(OutStr(3))
        TextADU.Text = OutStr(4): TextADL.Text = OutStr(5)
        TextDU.Text = OutStr(6): TextDL.Text = OutStr(7)
        TextUnit.Text = OutStr(8): TextFmtS.Text = OutStr(9)
        TextAlmU.Text = OutStr(10): TextAlmL.Text = OutStr(11)
        TextSFU.Text = OutStr(12): TextSFL.Text = OutStr(13)
        If Val(OutStr(14)) = 0 Then
            CheckCt.Value = 0
        Else
            CheckCt.Value = 1
        End If
        ComboCutImg.Text = OutStr(15)
        TextName.Text = OutStr(16)
    End If
    If UBound(OutStr) = 17 Then
        LabelNo.Caption = Val(OutStr(0)) & "-" & Format(Val(OutStr(1)), "00")
        ComboDelay.ListIndex = Val(OutStr(2)): ComboG.ListIndex = Val(OutStr(3))
        TextADU.Text = OutStr(4): TextADC.Text = OutStr(5): TextADL.Text = OutStr(6)
        TextDU.Text = OutStr(7): TextDL.Text = OutStr(8)
        TextUnit.Text = OutStr(9): TextFmtS.Text = OutStr(10)
        TextAlmU.Text = OutStr(11): TextAlmL.Text = OutStr(12)
        TextSFU.Text = OutStr(13): TextSFL.Text = OutStr(14)
        If Val(OutStr(15)) = 0 Then
            CheckCt.Value = 0
        Else
            CheckCt.Value = 1
        End If
        ComboCutImg.Text = OutStr(16)
        TextName.Text = OutStr(17)
    End If
End Sub

Private Sub TextPassW_Change()
    If TextPassW.Text = Password Then
        ComSave.Enabled = True
    Else
        ComSave.Enabled = False
    End If
End Sub

Private Sub TimerAD_Timer()
    TextAD.Text = MoniData(Val(Left(LabelNo.Caption, 1)), Val(Right(LabelNo.Caption, 2))).AD
End Sub
