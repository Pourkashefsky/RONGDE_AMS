Private Sub ComACK_Click()
Dim Ty As String, AX As Integer, AY As Integer
Dim S As String, Wh As Integer
    If SONK = True Then
        FrameErr.Visible = True
        Exit Sub
    End If
    Wh = ListAlm.ListIndex
    If Wh >= 0 And Wh <= ListAlm.ListCount - 1 Then
        S = ListAdd.List(Wh)
        Ty = Left(S, 1)
        S = Mid(S, 2): AX = Val(S)
        S = Mid(S, 4): AY = Val(S)
        Call FormMain.AckOne(Ty, AX, AY)
        Call ListAlmReport
    End If
    If Wh >= 0 And Wh <= ListAlm.ListCount - 1 Then
        ListAlm.ListIndex = Wh
    End If
End Sub

Private Sub ComACKPage_Click()
Dim Ty As String, AX As Integer, AY As Integer
Dim S As String
Dim i As Integer
    If SONK = True Then
        FrameErr.Visible = True
        Exit Sub
    End If
    For i = 0 To ListAlm.ListCount - 1
        S = ListAdd.List(i)
        Ty = Left(S, 1)
        S = Mid(S, 2): AX = Val(S)
        S = Mid(S, 4): AY = Val(S)
        Call FormMain.AckOne(Ty, AX, AY)
    Next i
    Call ListAlmReport
End Sub

Private Sub ComSilence_Click()
    Call FormMain.MnuUserSilence_Click
    FrameErr.Visible = False
End Sub

Private Sub Timer1_Timer()
    ImageSONK.Visible = SONK
End Sub
