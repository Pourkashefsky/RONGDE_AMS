Public Sub ComOK_Click()
    '检查正确的密码
    Watch1 = ComboWatch1.ListIndex    '值班人设置
    Watch2 = ComboWatch2.ListIndex
    DealyW1 = ComboDealy1.ListIndex
    DealyW2 = ComboDealy2.ListIndex
    DealyW3 = ComboDealy3.ListIndex
    SaveSetting "RDMS System", "Duty", "Watch1", Watch1
    SaveSetting "RDMS System", "Duty", "Watch2", Watch2
    SaveSetting "RDMS System", "Duty", "DealyW1", DealyW1
    SaveSetting "RDMS System", "Duty", "DealyW2", DealyW2
    SaveSetting "RDMS System", "Duty", "DealyW3", DealyW3
    Unload Me
End Sub

Public Sub Form_Load()
Dim i As Integer, j As Integer, k As Integer
Dim x As Integer, y As Integer, ImgErr As Boolean
    
    For i = 0 To 4
        ComboWatch1.AddItem WatchName(i)
        ComboWatch2.AddItem WatchName(i)
    Next i
    ComboWatch1.ListIndex = Watch1
    ComboWatch2.ListIndex = Watch2
    
    ComboDealy1.Clear
    ComboDealy1.AddItem "0 minute"
    ComboDealy1.AddItem "1 minute"
    ComboDealy1.AddItem "2 minute"
    ComboDealy1.AddItem "5 minute"
    ComboDealy1.AddItem "10 minute"
    ComboDealy1.ListIndex = DealyW1
    ComboDealy2.Clear
    ComboDealy2.AddItem "0 minute"
    ComboDealy2.AddItem "1 minute"
    ComboDealy2.AddItem "2 minute"
    ComboDealy2.AddItem "5 minute"
    ComboDealy2.AddItem "10 minute"
    ComboDealy2.ListIndex = DealyW2
    ComboDealy3.Clear
    ComboDealy3.AddItem "0 minute"
    ComboDealy3.AddItem "1 minute"
    ComboDealy3.AddItem "2 minute"
    ComboDealy3.AddItem "5 minute"
    ComboDealy3.AddItem "10 minute"
    ComboDealy3.ListIndex = DealyW3
    
End Sub

Private Sub txtPassword_Change()
    If txtPassword.Text = Password Then
        ComOK.Enabled = True
    Else
        ComOK.Enabled = False
    End If
End Sub
