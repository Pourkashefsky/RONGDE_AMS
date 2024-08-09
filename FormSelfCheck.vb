Private Sub Form_Load()
Dim N As Integer, i As Integer
    For N = BinCallSta To BinCallEnd
        TextTYK(N).Visible = True
    Next N
    For N = MoniCallSta To MoniCallEnd
        TextTYM(N).Visible = True
    Next N
    If MeMain = True Then
        For N = EMRCallSta To EMRCallEnd
            TextTYE(N).Visible = True
        Next N
    End If
    If FormMain.MSCommKQ16.CommPort <> 16 Then TextTYH(0).Visible = False
    If FormMain.MSCommQ32.CommPort <> 16 Then TextTYH(1).Visible = False
    If FormMain.MSCommU24.CommPort <> 16 Then TextTYH(2).Visible = False
    
End Sub

Private Sub Timer1_Timer()
Dim N As Integer, S As String
    For N = 0 To 9
        If BinCall(N).CallFail > BinCall(N).MaxCall Then
            TextTYK(N).Text = "RDDS-K32   " & N & "#" & vbTab & "ERR!!!" & " D:" & BinCall(N).CallFail
            TextTYK(N).BackColor = vbRed
        Else
            TextTYK(N).Text = "RDDS-K32   " & N & "#" & vbTab & "OK!!!" & " D:" & BinCall(N).CallFail
            TextTYK(N).BackColor = vbGreen
        End If
        If MoniCall(N).CallFail > MoniCall(N).MaxCall Then
            TextTYM(N).Text = "RDDS-A/T24 " & N & "#" & vbTab & "ERR!!!" & " D:" & MoniCall(N).CallFail
            TextTYM(N).BackColor = vbRed
        Else
            TextTYM(N).Text = "RDDS-A/T24 " & N & "#" & vbTab & "OK!!!" & " D:" & MoniCall(N).CallFail
            TextTYM(N).BackColor = vbGreen
        End If
        If EMRCall(N).CallFail > EMRCall(N).MaxCall Then
            TextTYE(N).Text = "RDDS-EAP " & EMRNo(N) & "#" & vbTab & "ERR!!!" & " D:" & EMRCall(N).CallFail
            TextTYE(N).BackColor = vbRed
        Else
            TextTYE(N).Text = "RDDS-EAP " & EMRNo(N) & "#" & vbTab & "OK!!!" & " D:" & EMRCall(N).CallFail
            TextTYE(N).BackColor = vbGreen
        End If
    Next N
    If KQCall.CallFail > KQCall.MaxCall Then
        TextTYH(0).Text = "RDDS-KQ16  0#" & vbTab & "ERR!!!" & " D:" & KQCall.CallFail
        TextTYH(0).BackColor = vbRed
    Else
        TextTYH(0).Text = "RDDS-KQ16  0#" & vbTab & "OK!!!" & " D:" & KQCall.CallFail
        TextTYH(0).BackColor = vbGreen
    End If
    If Q32Call.CallFail > Q32Call.MaxCall Then
        TextTYH(1).Text = "RDDS-Q32   0#" & vbTab & "ERR!!!" & " D:" & Q32Call.CallFail
        TextTYH(1).BackColor = vbRed
    Else
        TextTYH(1).Text = "RDDS-Q32   0#" & vbTab & "OK!!!" & " D:" & Q32Call.CallFail
        TextTYH(1).BackColor = vbGreen
    End If
    If U24Call.CallFail > U24Call.MaxCall Then
        TextTYH(2).Text = "RDDS-U24   0#" & vbTab & "ERR!!!" & " D:" & U24Call.CallFail
        TextTYH(2).BackColor = vbRed
    Else
        TextTYH(2).Text = "RDDS-U24   0#" & vbTab & "OK!!!" & " D:" & U24Call.CallFail
        TextTYH(2).BackColor = vbGreen
    End If
    If MS1Call.CallFail > MS1Call.MaxCall Then
        TextTYH(3).Text = "No1 RS232" & vbTab & "ERR!!!"
        TextTYH(3).BackColor = vbRed
    Else
        TextTYH(3).Text = "No1 RS232" & vbTab & "OK!!!"
        TextTYH(3).BackColor = vbGreen
    End If
    If MS2Call.CallFail > MS2Call.MaxCall Then
        TextTYH(4).Text = "No2 RS232" & vbTab & "ERR!!!"
        TextTYH(4).BackColor = vbRed
    Else
        TextTYH(4).Text = "No2 RS232" & vbTab & "OK!!!"
        TextTYH(4).BackColor = vbGreen
    End If
    
    If MeMain = False Then Unload Me
End Sub
