Private Sub ComEdit_Click()
    List1.List(List1.ListIndex) = Left(LabelNo.Caption, 1) & vbTab & _
        Right(LabelNo.Caption, 2) & vbTab & Format(ComboDelay.ListIndex, "00") & vbTab & _
        ComboG.ListIndex & vbTab & CheckNor.Value & vbTab & CheckCt.Value & vbTab & ComboCutImg.Text & vbTab & TextName.Text
End Sub

Private Sub ComSave_Click()
    Call List2File(List1, App.Path & "\SetB.ini", LabTital.Caption)
    Call LoadSet
    Unload Me
End Sub

Private Sub Form_Load()
Dim N As Integer, i As Integer, Nr As Integer, Ct As Integer
    For i = 0 To Maxdelay
        ComboDelay.AddItem Format(i * 3.3, "00.0 S")
    Next i
    For i = 0 To 9
        ComboG.AddItem i
    Next i
    
    For N = 0 To 9
        For i = 0 To 31
            If BinData(N, i).Nor = True Then
                Nr = 1
            Else
                Nr = 0
            End If
            If BinData(N, i).Cutout = True Then
                Ct = 1
            Else
                Ct = 0
            End If
            
            List1.AddItem N & vbTab & Format(i, "00") & vbTab & _
                Format(BinData(N, i).Delay, "00") & vbTab & BinData(N, i).Group & vbTab & _
                Nr & vbTab & Ct & vbTab & BinData(N, i).CutImg & vbTab & BinData(N, i).Name
            
            ComboCutImg.AddItem N & "-" & Format(i, "00")
        Next i
    Next N
    ComboCutImg.AddItem "9-99"
    List1.ListIndex = 0
    
End Sub

Private Sub List1_Click()
    Call Str2Array(List1.Text)
    If UBound(OutStr) = 7 Then
        LabelNo.Caption = Val(OutStr(0)) & "-" & Format(Val(OutStr(1)), "00")
        ComboDelay.ListIndex = Val(OutStr(2))
        ComboG.ListIndex = Val(OutStr(3))
        If Val(OutStr(4)) = 0 Then
            CheckNor.Value = 0
        Else
            CheckNor.Value = 1
        End If
        If Val(OutStr(5)) = 0 Then
            CheckCt.Value = 0
        Else
            CheckCt.Value = 1
        End If
        ComboCutImg.Text = OutStr(6)
        TextName.Text = OutStr(7)
    End If
End Sub

Private Sub TextPassW_Change()
    If TextPassW.Text = Password Then
        ComSave.Enabled = True
    Else
        ComSave.Enabled = False
    End If
End Sub
