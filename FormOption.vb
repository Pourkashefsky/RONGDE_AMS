Private Sub ComCOMM_Click()
Dim i As Integer, j As Integer
With FormMain
    For i = 1 To 8
        .PicGroup(i).Visible = False
    Next i
    .PicK32.Visible = False
    .PicAT24.Visible = False
    .PicComm.Visible = True
    
End With
    Unload Me
End Sub

Private Sub ComK32_Click()
Dim i As Integer
With FormMain
    For i = 1 To 8
        .PicGroup(i).Visible = False
    Next i
    .PicK32.Visible = True
    .PicAT24.Visible = False
    .PicComm.Visible = False
End With
    Unload Me
End Sub

Private Sub ComKSet_Click()
    SetupK32.Show
    Unload Me
End Sub

Private Sub ComM24_Click()
Dim i As Integer
With FormMain
    For i = 1 To 8
        .PicGroup(i).Visible = False
    Next i
    .PicK32.Visible = False
    .PicAT24.Visible = True
    .PicComm.Visible = False
End With
    Unload Me
End Sub

Private Sub ComMSet_Click()
    SetupM24.Show
    Unload Me
End Sub

Private Sub ComU24_Click()
    SetupU24.Show
    Unload Me
End Sub

Private Sub ComVDR_Click()
    SetupVDR.Show
    Unload Me
End Sub

Private Sub TextPassW_Change()
Dim S As String
    S = TextPassW.Text
    If S = "rongded" Or S = "RONGDED" Then
        Me.Height = 3270
    Else
        If S = Password And MeMain = True Then
            Me.Height = 1845
        Else
            Me.Height = 1050
        End If
    End If
End Sub

Private Sub TextPassW_GotFocus()
    TextPassW.SelStart = 0
    TextPassW.SelLength = Len(TextPassW.Text)
End Sub
