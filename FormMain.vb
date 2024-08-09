Private Sub CheckAutoPrint_Click()
    If CheckAutoPrint.Value = 0 Then
        AutoPrint = False
    Else
        AutoPrint = True
    End If
    SaveSetting "RDMS", "OPT", "AutoPrint", CheckAutoPrint.Value
End Sub

Private Sub CheckAutoZorder_Click()
    If CheckAutoZorder.Value = 0 Then
        AutoZorder = False
    Else
        AutoZorder = True
    End If
    SaveSetting "RDMS", "OPT", "AutoZorder", CheckAutoZorder.Value
End Sub

Private Sub CheckMain_Click()
    SaveSetting "RDMS", "OPT", "Main", CheckMain.Value
End Sub

Public Sub ComACK_Click()
    Call MnuUserACK_Click
    Unload FormReport
End Sub

Private Sub ComCTime_Click()
    FormTime.Show
End Sub

Private Sub ComDuty_Click()
    Call MnuLogin_Click
End Sub

Private Sub ComGroup_Click(Index As Integer)
    MnuGroup_Click (Index)
End Sub

Private Sub ComOption_Click()
    Call MnuOption_Click
End Sub

Private Sub ComRecord_Click()
    Call MnuRecord_Click
End Sub

Private Sub ComReport_Click()
    Call ListAlmReport
End Sub

Private Sub ComSilence_Click()
    Call MnuUserSilence_Click
End Sub

Private Sub ComSysAlm_Click()
    FormSelfCheck.Show
End Sub

Private Sub Form_Initialize()
    Call AlmListClr     '报警列表堆栈初始化
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
    CheckMain.Value = GetSetting("RDMS", "OPT", "Main", 0)      '默认为主机/从机
    
    MeMain = False
    Me.Move 0, 0, 15360, 11520
    Call LoadSet
    Call ComGroup_Click(1)
    If GetSetting("RDMS", "OPT", "AutoPrint", 1) = 1 Then
        AutoPrint = True
        CheckAutoPrint.Value = 1
    Else
        AutoPrint = False
        CheckAutoPrint.Value = 0
    End If
    If GetSetting("RDMS", "OPT", "AutoZorder", 1) = 1 Then
        AutoZorder = True
        CheckAutoZorder.Value = 1
    Else
        AutoZorder = False
        CheckAutoZorder.Value = 0
    End If
    
    For i = 0 To 4
        ComboNT(i).Clear
        ComboNV(i).Clear
        For j = 0 To 9
            ComboNT(i).AddItem j
            ComboNV(i).AddItem j
        Next j
        ComboNT(i).ListIndex = 0
        ComboNV(i).ListIndex = 0
        
        ComboI(i).Clear
        For j = 0 To 31
            ComboI(i).AddItem j
        Next j
        ComboI(i).ListIndex = 0
        
        ComboIV(i).Clear
        For j = 0 To 23
            ComboIV(i).AddItem j
        Next j
        ComboIV(i).ListIndex = 0
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPassW.Show
    Cancel = 1
End Sub

Private Sub MnuGroup_Click(Index As Integer)
Dim i As Integer
    PicGroup(Index).Visible = True
    PicGroup(Index).ZOrder
    For i = 1 To 8
        If i <> Index Then PicGroup(i).Visible = False
    Next i
    PageNow = Index
End Sub

Private Sub MnuLogin_Click()
    FormDuty.Show
End Sub

Private Sub MnuOption_Click()
    FormOption.Height = 1050
    FormOption.Show
End Sub

Private Sub MnuRecord_Click()
    FormRecord.Show
End Sub

Public Sub MnuUserACK_Click()
On Error Resume Next
Dim i As Integer, N As Integer, Prn As String
Dim PN As Integer
Dim AX As Integer, AY As Integer

    If SONK = True Then Exit Sub    '先消声再消闪
    
    Open "prn" For Output As #1
    
    If PageNow <> 8 Then    '非平均温差报警界面
        For PN = 0 To ListNum - 1
            AX = PxAL(PageNow, PN).AddX
            AY = PxAL(PageNow, PN).AddY
            Select Case PxAL(PageNow, PN).BorM
            Case "B"
                i = AX * 32 + AY
                Prn = ""
                Prn = BinData(AX, AY).Name
                
                If LabCHName(i).ToolTipText = "ALM" Then
                    LabCHName(i).ToolTipText = "ACKN-ALM"
                    ShapeAlm(i).FillColor = vbRed
                    If Prn <> "" And TimerSaveEnable.Enabled = False Then
                        Call AlmListAdd("B", AX, AY, "A")
                        '...............存入数据库
                        FormRecord.Adodc1.Recordset.AddNew
                        FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                        FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                        FormRecord.Adodc1.Recordset("Value") = "ACKN-ALM"
                        FormRecord.Adodc1.Recordset.Update
                        Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-ALM"
                        If MeMain = True Then
                            If AutoPrint = True Then Print #1, Prn
                            ListSlaveSave.AddItem Prn
                        End If
                    End If
                End If
                If LabCHName(i).ToolTipText = "SF" Then
                    LabCHName(i).ToolTipText = "ACKN-SF"
                    ShapeAlm(i).FillColor = vbYellow
                    If Prn <> "" And TimerSaveEnable.Enabled = False Then
                        '...............存入数据库
                        FormRecord.Adodc1.Recordset.AddNew
                        FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                        FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                        FormRecord.Adodc1.Recordset("Value") = "ACKN-SF"
                        FormRecord.Adodc1.Recordset.Update
                        Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-SF"
                        If MeMain = True Then
                            If AutoPrint = True Then Print #1, Prn
                            ListSlaveSave.AddItem Prn
                        End If
                    End If
                End If
            Case "M"
                i = AX * 24 + AY
                Prn = ""
                Prn = MoniData(AX, AY).Name
                
                If LabCHNameM(i).ToolTipText = "ALM" Then
                    LabCHNameM(i).ToolTipText = "ACKN-ALM"
                    ShapeAlmM(i).FillColor = vbRed
                    If Prn <> "" And TimerSaveEnable.Enabled = False Then
                        Call AlmListAdd("M", AX, AY, "A")
                        '...............存入数据库
                        FormRecord.Adodc1.Recordset.AddNew
                        FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                        FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                        FormRecord.Adodc1.Recordset("Value") = "ACKN-ALM"
                        FormRecord.Adodc1.Recordset.Update
                        Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-ALM"
                        If MeMain = True Then
                            If AutoPrint = True Then Print #1, Prn
                            ListSlaveSave.AddItem Prn
                        End If
                    End If
                End If
                
                If LabCHNameM(i).ToolTipText = "SF" Then
                    LabCHNameM(i).ToolTipText = "ACKN-SF"
                    ShapeAlmM(i).FillColor = vbYellow
                    If Prn <> "" And TimerSaveEnable.Enabled = False Then
                        '...............存入数据库
                        FormRecord.Adodc1.Recordset.AddNew
                        FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                        FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                        FormRecord.Adodc1.Recordset("Value") = "ACKN-SF"
                        FormRecord.Adodc1.Recordset.Update
                        Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-SF"
                        If MeMain = True Then
                            If AutoPrint = True Then Print #1, Prn
                            ListSlaveSave.AddItem Prn
                        End If
                    End If
                End If
            End Select
        Next PN
    Else
        N = 0
        AX = 2      '温度报警确认
        For AY = 6 To 14
            i = AX * 24 + AY
            Prn = MoniData(AX, AY).Name
            
            If LabCHNameM(i).ToolTipText = "ALM" Then
                LabCHNameM(i).ToolTipText = "ACKN-ALM"
                ShapeAlmM(i).FillColor = vbRed
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    Call AlmListAdd("M", AX, AY, "A")
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-ALM"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-ALM"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
            If LabCHNameM(i).ToolTipText = "SF" Then
                LabCHNameM(i).ToolTipText = "ACKN-SF"
                ShapeAlmM(i).FillColor = vbYellow
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-SF"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-SF"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
            '平均温差报警确认
            Prn = "DEVIATION ALARM " & N + 1 & "#"
            If LabTYP(N).ToolTipText = "ALM" Then
                LabTYP(N).ToolTipText = "ACKN-ALM"
                ShapeAlmPJ(N).FillColor = vbRed
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    Call AlmListAdd("M", AX, AY, "A")
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-ALM"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-ALM"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
            If LabTYP(N).ToolTipText = "SF" Then
                LabTYP(N).ToolTipText = "ACKN-SF"
                ShapeAlmPJ(N).FillColor = vbYellow
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-SF"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-SF"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
            N = N + 1
        Next AY
    End If
    Close #1
    DoEvents
End Sub

'Ty 类型 B开关量 M模拟量
'AX,AY 地址
Public Function AckOne(Ty As String, AX As Integer, AY As Integer)
Dim i As Integer, Prn As String
    If SONK = True Then Exit Function '先消声再消闪
    Open "prn" For Output As #1
    Select Case Ty
        Case "B"
            i = AX * 32 + AY
            Prn = ""
            Prn = BinData(AX, AY).Name
            
            If LabCHName(i).ToolTipText = "ALM" Then
                LabCHName(i).ToolTipText = "ACKN-ALM"
                ShapeAlm(i).FillColor = vbRed
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    Call AlmListAdd("B", AX, AY, "A")
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-ALM"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-ALM"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
            If LabCHName(i).ToolTipText = "SF" Then
                LabCHName(i).ToolTipText = "ACKN-SF"
                ShapeAlm(i).FillColor = vbYellow
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-SF"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-SF"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
        Case "M"
            i = AX * 24 + AY
            Prn = ""
            Prn = MoniData(AX, AY).Name
            
            If LabCHNameM(i).ToolTipText = "ALM" Then
                LabCHNameM(i).ToolTipText = "ACKN-ALM"
                ShapeAlmM(i).FillColor = vbRed
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    Call AlmListAdd("M", AX, AY, "A")
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-ALM"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-ALM"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
            If LabCHNameM(i).ToolTipText = "SF" Then
                LabCHNameM(i).ToolTipText = "ACKN-SF"
                ShapeAlmM(i).FillColor = vbYellow
                If Prn <> "" And TimerSaveEnable.Enabled = False Then
                    '...............存入数据库
                    FormRecord.Adodc1.Recordset.AddNew
                    FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
                    FormRecord.Adodc1.Recordset("Name") = Left(Prn, 50)
                    FormRecord.Adodc1.Recordset("Value") = "ACKN-SF"
                    FormRecord.Adodc1.Recordset.Update
                    Prn = Format(Now, tFmsL) & vbTab & Prn & vbTab & "ACKN-SF"
                    If MeMain = True Then
                        If AutoPrint = True Then Print #1, Prn
                        ListSlaveSave.AddItem Prn
                    End If
                End If
            End If
    End Select
    Close #1
    DoEvents
End Function

Public Sub MnuUserSilence_Click()
    SONK = False
    SONKJX = False
End Sub

Private Sub TimerAT24_Timer()
Dim S As String, i As Integer, j As Integer, k As Integer
Dim Rd() As Byte, Rds() As Byte, OK As Boolean  '定义字符串存放数组，非乱码的标志
Dim N As Integer, S1 As String, INS() As Byte, S2 As String
Dim AD As Integer, ADL As Integer, ADU As Integer
Dim DispU As Single, DispL As Single, Disp As Single, Sum As Double
Dim x As Integer, y As Integer, z As Integer
Dim MC As Integer
    If MeMain = False Then Exit Sub
    If MSCommAT24.CommPort = 16 Then Exit Sub
    
    '如果是主机则呼叫一次AT24
    N = MoniCallNo
    S = "#I" & N & "M$"
    If MSCommAT24.PortOpen = True And MeMain = True Then
        MSCommAT24.Output = S
        LabCallM.Caption = S
    End If
    MoniCall(N).CallFail = MoniCall(N).CallFail + 1
    MoniCallNo = MoniCallNo + 1
    If MoniCallNo > MoniCallEnd Then MoniCallNo = MoniCallSta
        
    S = ""
    If MSCommAT24.PortOpen = True Then
        '取出后滤波并放到尾部
        Rd = MSCommAT24.Input
        If UBound(Rd) >= 0 Then
            OK = True
            ReDim Rds(UBound(Rd) * 2) As Byte
            For i = 0 To UBound(Rd)
                If (Rd(i) > 126 Or Rd(i) < 32) And Rd(i) <> 10 And Rd(i) <> 13 Then
                    OK = False
                End If
                Rds(i * 2) = Rd(i)
            Next i
            If OK = True Then
                S = Rds
            End If
        End If
        RTReadM.SelStart = Len(RTReadM.Text)
        RTReadM.SelText = S          '文本显示
        
        If TxtRead(RTReadM, "*", "$", 103) = True Then    '处理数据
            S = OutStr(0): LabGetM.Caption = DealS
            N = Val(Mid(S, 2, 1))
            MoniCall(N).CallFail = 0     '通讯成功
            If UBound(OutStr) = 24 Then
            For i = 0 To 23
                S1 = Right(OutStr(i), 3)
                INS = S1
                x = INS(0) - 48
                y = INS(2) - 48
                z = INS(4) - 48
                AD = x * 256 + y * 16 + z
                If AD <= MoniData(N, i).ADC Then
                    ADL = MoniData(N, i).ADL: ADU = MoniData(N, i).ADC
                    DispL = MoniData(N, i).DispL
                    DispU = (MoniData(N, i).DispU + MoniData(N, i).DispL) / 2
                Else
                    ADL = MoniData(N, i).ADC: ADU = MoniData(N, i).ADU
                    DispL = (MoniData(N, i).DispU + MoniData(N, i).DispL) / 2
                    DispU = MoniData(N, i).DispU
                End If
                
                MoniData(N, i).AD = AD
                If ADU - ADL = 0 Then
                    Disp = 0
                Else
                    Disp = (AD - ADL) * (DispU - DispL) / (ADU - ADL) + DispL
                    If Disp > DispU Then Disp = DispU
                    If Disp < DispL Then Disp = DispL
                End If
                
                '数组前移一位
                For k = 0 To MoniData(N, i).Delay - 1
                    MoniData(N, i).DelayV(k) = MoniData(N, i).DelayV(k + 1)
                Next k
                '新数据放在数组最后
                MoniData(N, i).DelayV(MoniData(N, i).Delay) = Disp
                '计算显示值(数组平均数)
                Sum = 0
                For k = 0 To MoniData(N, i).Delay
                    Sum = Sum + MoniData(N, i).DelayV(k)
                Next k
                S2 = Sum / (MoniData(N, i).Delay + 1)
                MoniData(N, i).Value = Val(Format(S2, MoniData(N, i).FmtS))
                
                '判断报警状态
                If MoniData(N, i).Value >= MoniData(N, i).AlmU Or MoniData(N, i).Value <= MoniData(N, i).AlmL Then
                    MoniData(N, i).Alm = True
                End If
                If MoniData(N, i).Value < MoniData(N, i).AlmU * 0.98 And MoniData(N, i).Value > MoniData(N, i).AlmL * 1.02 Then
                    MoniData(N, i).Alm = False
                End If
                If MoniData(N, i).AD >= MoniData(N, i).SFU Or MoniData(N, i).AD <= MoniData(N, i).SFL Then
                    MoniData(N, i).SF = True
                Else
                    MoniData(N, i).SF = False
                End If
            Next i
            End If
        End If
    End If
    
    '通讯故障判断
    For i = 0 To 9
        If MoniCall(i).CallFail >= MoniCall(i).MaxCall Then
            MoniCall(i).CallFail = 999
            If MoniCall(i).CommFail = False Then
                MoniCall(i).CommFail = True
                FormSelfCheck.Show
                SONK = True
            End If
            For j = 0 To 23
                MoniData(i, j).SF = True
            Next j
        Else
            MoniCall(i).CommFail = False
        End If
    Next i
    
    'TEST 强制置位
    For i = 0 To 4
        If CheckTestV(i).Value = 1 Then
            MoniData(ComboNV(i).ListIndex, ComboIV(i).ListIndex).Value = Val(TextValue(i).Text)
            MoniData(ComboNV(i).ListIndex, ComboIV(i).ListIndex).Alm = CBool(CheckAlmV(i).Value)
            MoniData(ComboNV(i).ListIndex, ComboIV(i).ListIndex).SF = CBool(CheckSFV(i).Value)
        End If
    Next i
End Sub

Private Sub TimerDeal_Timer()
Dim RunTime As Long
    RunTime = TimerDeal.Interval * Val(TimerDeal.Tag)
    If RunTime Mod TimerMS.Interval = 0 Then
        Call TimerMS_Timer
    End If
    If RunTime Mod TimerK32.Interval = 0 Then
        Call TimerK32_Timer
    End If
    If RunTime Mod TimerAT24.Interval = 0 Then
        Call TimerAT24_Timer
    End If
    If RunTime Mod TimerKQ16.Interval = 0 Then
        Call TimerKQ16_Timer
    End If
    If RunTime Mod TimerQ32.Interval = 0 Then
        Call TimerQ32_Timer
    End If
    If RunTime Mod TimerEMR.Interval = 0 Then
        Call TimerEMR_Timer
    End If
    If RunTime Mod TimerU24.Interval = 0 Then
        Call TimerU24_Timer
    End If
    If RunTime Mod TimerVDR.Interval = 0 Then
        Call TimerVDR_Timer
    End If
    If RunTime Mod TimerReset.Interval = 0 Then
        Call TimerReset_Timer
    End If
    
    TimerDeal.Tag = Val(TimerDeal.Tag) + 1
    If Val(TimerDeal.Tag) >= 100 Then TimerDeal.Tag = "0"
End Sub

Private Sub TimerEMR_Timer()
Dim N As Integer, i As Integer
Dim SM As String, SG As String, SA As String, ST As String, SC As String, SR As String
Dim Stmp As String
Dim CTorF As Integer        '该台是否鸣叫 1 - 叫
Dim CallDealy As Integer    'Login 单项设置时间间隔(S)
Dim DW1 As Integer, DW2 As Integer, DW3 As Integer  '各延时长度（秒）
Dim BackN As Integer, AX As Integer '应答台号
Dim QTF As Boolean
Dim Pi As Integer, x As Integer, y As Integer
    If MeMain = False Then Exit Sub
    If MSCommEMR.CommPort = 16 Then Exit Sub
    If MSCommEMR.PortOpen = False Then Exit Sub
    If MSCommEMR.OutBufferCount <> 0 Then Exit Sub
    
    EMRNo(0) = 7    '机舱       地址7
    EMRNo(1) = 0    '老轨       地址0
    EMRNo(2) = 1    '大轨       地址1
    EMRNo(3) = 2    '二轨       地址2
    EMRNo(4) = 3    '驾驶室     地址3
    EMRNo(5) = 8    '公共房间   地址8   货控室
    EMRNo(6) = 9    '公共房间   地址9   吸烟室
    EMRNo(7) = 10   '公共房间   地址10  餐厅
    EMRNo(8) = 11   '公共房间   地址11  会议室
    EMRNo(9) = 99   '备用
    
    Select Case DealyW1
    Case 0  '0 minute"
        DW1 = 0
    Case 1  '1 minute"
        DW1 = 60
    Case 2  '2 minute"
        DW1 = 120
    Case 3  '5 minute"
        DW1 = 300
    Case 4  '10 minute"
        DW1 = 600
    End Select
    Select Case DealyW2
    Case 0  '0 minute"
        DW2 = 0
    Case 1  '1 minute"
        DW2 = 60
    Case 2  '2 minute"
        DW2 = 120
    Case 3  '5 minute"
        DW2 = 300
    Case 4  '10 minute"
        DW2 = 600
    End Select
    Select Case DealyW3
    Case 0  '0 minute"
        DW3 = 0
    Case 1  '1 minute"
        DW3 = 60
    Case 2  '2 minute"
        DW3 = 120
    Case 3  '5 minute"
        DW3 = 300
    Case 4  '10 minute"
        DW3 = 600
    End Select
    
    If TimeEAP = False Then     '如时间未同步过则同步时间
        Stmp = "T" & Format(Now, "yy-mm-dd hh:mm:ss")
        ST = Left(Stmp, 3) & "/" & Mid(Stmp, 5, 2) & "/" & Mid(Stmp, 8, 2) & "/" & Mid(Stmp, 11, 2) & "/" & Mid(Stmp, 14, 2) & "/" & Mid(Stmp, 17, 2)
        ST = "#" & ST & "$" & ASCSumF(ST, 2) & vbCrLf
        RTEMR.Text = Now & vbTab & "Indx=" & EMRTmN1 & "/" & EMRTmN2 & vbCrLf & "Send:" & vbCrLf & ST
        MSCommEMR.Output = ST
        TimeEAP = True
        QTF = True
    End If
    If QTF = True Then Exit Sub '同步之后延迟一下
    
    Select Case EMRTmN1
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9   '各台呼叫
        'EMRCallFail(9)     '延伸面板呼叫失败次数
        'Public MaxEMRCallFail As Integer    '延伸面板呼叫失败最大次数
        'CTorF As Integer    '该台是否鸣叫 1-叫
        N = EMRTmN1
        CTorF = 0
        If timeSONK > DW1 Then        '值班人报警
            If Watch1 <= 3 And EMRNo(N) = Watch1 Then CTorF = 1    '有人值班而且呼叫号为值班人号
        End If
        If timeSONK > (DW1 + DW2) Then   '值班长报警
            If Watch2 <= 3 And EMRNo(N) = Watch2 Then CTorF = 1    '有人值班而且呼叫号为值班人号
        End If
        If timeSONK > (DW1 + DW2 + DW3) Then  '全部报警
            CTorF = 1
        End If
        
        SC = "H" & Format(EMRNo(N), "00") & CTorF
        SC = "#" & SC & "$" & ASCSumF(SC, 2) & vbCrLf
        RTEMR.Text = Now & vbTab & "Indx=" & EMRTmN1 & "/" & EMRTmN2 & vbCrLf & "Send:" & vbCrLf & SC
        MSCommEMR.Output = SC   '#HXXY$hh[0DH][0AH]
        DoEvents
        EMRCall(N).CallFail = EMRCall(N).CallFail + 1
        
        Stmp = ""
        Stmp = MSCommEMR.Input  '*XXY$[0DH][0AH]   读串口应答/置通讯故障判断
        RTEMR.Text = RTEMR.Text & vbCrLf & "Read:" & vbCrLf & Stmp
        If Left(Stmp, 1) = "*" And Mid(Stmp, 5, 1) = "$" Then
            BackN = Val(Mid(Stmp, 2, 2))
            AX = 0
            Do Until EMRNo(AX) = BackN
                AX = AX + 1
            Loop
            EMRCall(AX).CallFail = 0
            If Mid(Stmp, 4, 1) = "1" Then   '对方确认过
                timeSONK = 0
                EMRTmN1 = BackN
            End If
        End If
    Case 10  'Moni量延伸 显示通讯
        SM = "M" & AlmListStr
        
        SM = "#" & SM & "$" & ASCSumF(SM, 2) & vbCrLf
        RTEMR.Text = Now & vbTab & "Indx=" & EMRTmN1 & "/" & EMRTmN2 & vbCrLf & "Send:" & vbCrLf & SM
        MSCommEMR.Output = SM
        DoEvents
    Case 14  '组报警/值班人 显示通讯 0组不输出
        Call GroupDeal
        For i = 1 To 8
            SG = SG & Bin2Asc(False, False, GroupAlm(i).ExtALm, GroupAlm(i).NewAlm)
        Next i
        SG = "N" & Watch1 & SG
        SG = "#" & SG & "$" & ASCSumF(SG, 2) & vbCrLf
        RTEMR.Text = Now & vbTab & "Indx=" & EMRTmN1 & "/" & EMRTmN2 & vbCrLf & "Send:" & vbCrLf & SG
        MSCommEMR.Output = SG
        DoEvents
    Case Else
        RTEMR.Text = Now & vbTab & "Indx=" & EMRTmN1 & "/" & EMRTmN2
        DoEvents
    End Select
    
    'KQ16 通讯故障判定
    For i = 0 To 8
        If EMRCall(i).CallFail >= EMRCall(i).MaxCall Then
            EMRCall(i).CallFail = 999
            If EMRCall(i).CommFail = False Then
                EMRCall(i).CommFail = True
                FormSelfCheck.Show
                SONK = True
            End If
        Else
            EMRCall(i).CommFail = False
        End If
    Next i
    
    EMRTmN1 = EMRTmN1 + 1
    If EMRTmN1 > 15 Then
        EMRTmN1 = 0    '定时器小循环计数
        EMRTmN2 = EMRTmN2 + 1
    End If
End Sub

Private Sub TimerK32_Timer()
Dim S As String, i As Integer, j As Integer, k As Integer
Dim Rd() As Byte, Rds() As Byte, OK As Boolean  '定义字符串存放数组，非乱码的标志
Dim N As Integer, S1 As String, INS() As Byte, S2 As String
Dim AD As Integer, ADL As Integer, ADU As Integer
Dim DispU As Single, DispL As Single, Disp As Single, Sum As Double
Dim x As Integer, y As Integer, z As Integer
Dim MC As Integer
    If MeMain = False Then Exit Sub
    If MSCommK32.CommPort = 16 Then Exit Sub
    
    '如果是主机则呼叫一次K32
    N = BinCallNo
    S = "#I" & N & "N$"                     '循环呼叫
    If MeMain = True And MSCommK32.PortOpen = True Then
        MSCommK32.Output = S
        LabCallK.Caption = S
    End If
    BinCall(N).CallFail = BinCall(N).CallFail + 1
    BinCallNo = BinCallNo + 1
    If BinCallNo > BinCallEnd Then BinCallNo = BinCallSta
    
    S = ""
    If MSCommK32.PortOpen = True Then
        '取出后滤波并放到尾部
        Rd = MSCommK32.Input        '*I0N????????[ODHOAH]03$
        If UBound(Rd) >= 0 Then
            OK = True
            ReDim Rds(UBound(Rd) * 2) As Byte
            For i = 0 To UBound(Rd)
                If (Rd(i) > 126 Or Rd(i) < 32) And Rd(i) <> 10 And Rd(i) <> 13 Then
                    OK = False
                End If
                Rds(i * 2) = Rd(i)
            Next i
            If OK = True Then
                S = Rds
            End If
        End If
        RTReadK.SelStart = Len(RTReadK.Text)
        RTReadK.SelText = S          '文本显示
        
        If TxtRead(RTReadK, "*", "$", 15) = True Then    '处理数据
            S = OutStr(0): LabGetK.Caption = S
            N = Val(Mid(S, 2, 1))           '台号
            BinCall(N).CallFail = 0         '通讯成功
            INS = S
            For i = 0 To 7
                Call Byte2Bin(INS(6 + i * 2))
                
                '#20080526-------------------------------------------------------------------------------
                If N = 0 And i = 0 Then     '第0台的第1字节(0-3通道)特殊处理，不来自采集板而是程序判定
                    Bin(0) = False: Bin(1) = False: Bin(2) = False: Bin(3) = False
                    If MoniData(2, 2).Value < 1.6 Then Bin(0) = True    'ME LO INLET PRESS LOW
                    If MoniData(1, 14).Value > 92 Then Bin(1) = True    'ME HT WATER JACKET OUTLET TEMP
                    If MoniData(2, 4).Value < 1.3 Then Bin(2) = True    'ME HT WATER JACKET INLET PRESS
                    If MoniData(2, 6).Value > 520 Or _
                        MoniData(2, 7).Value > 520 Or _
                        MoniData(2, 8).Value > 520 Or _
                        MoniData(2, 9).Value > 520 Or _
                        MoniData(2, 10).Value > 520 Or _
                        MoniData(2, 11).Value > 520 Or _
                        MoniData(2, 12).Value > 520 Or _
                        MoniData(2, 13).Value > 520 Or _
                        MoniData(2, 14).Value > 520 Then                'ME EXH GAS CYL 1-9 OUTLET TEMP
                        Bin(3) = True
                    End If
                End If
                '#20080526-------------------------------------------------------------------------------
                
                For j = 0 To 3
                    '数组前移一位
                    For k = 0 To BinData(N, i * 4 + j).Delay - 1
                        BinData(N, i * 4 + j).DelayA(k) = BinData(N, i * 4 + j).DelayA(k + 1)
                    Next k
                    '新数据放在数组最后
                    BinData(N, i * 4 + j).DelayA(BinData(N, i * 4 + j).Delay) = Bin(j)
                    '判断报警状态/正常优先
                    BinData(N, i * 4 + j).Alm = True
                    For k = 0 To BinData(N, i * 4 + j).Delay
                        If BinData(N, i * 4 + j).Nor = False Then
                            If BinData(N, i * 4 + j).DelayA(k) = False Then BinData(N, i * 4 + j).Alm = False
                        Else
                            If BinData(N, i * 4 + j).DelayA(k) = True Then BinData(N, i * 4 + j).Alm = False
                        End If
                    Next k
                Next j
            Next i
        End If
    End If
    
    '通讯故障判断
    For i = 0 To 9
        If BinCall(i).CallFail >= BinCall(i).MaxCall Then
            BinCall(i).CallFail = 999
            If BinCall(i).CommFail = False Then
                BinCall(i).CommFail = True
                FormSelfCheck.Show
                SONK = True
            End If
            For j = 0 To 31
                BinData(i, j).SF = True
            Next j
        Else
            BinCall(i).CommFail = False
            For j = 0 To 31
                BinData(i, j).SF = False
            Next j
        End If
    Next i
    
    'TEST 强制置位
    For i = 0 To 4
        If CheckTest(i).Value = 1 Then
            BinData(ComboNT(i).ListIndex, ComboI(i).ListIndex).Alm = CBool(CheckAlm(i).Value)
            BinData(ComboNT(i).ListIndex, ComboI(i).ListIndex).SF = CBool(CheckSF(i).Value)
        End If
    Next i
    
    '强制置位把KQ16点取出到K32
    'BinData(8, 10).Alm = K16(2)  'UPS电源故障
    'BinData(8, 9).Alm = K16(3) '死人报警
    
End Sub

Private Sub TimerKQ16_Timer()
Dim S1 As String, S2 As String, S3 As String, S4 As String
Dim S As String
Dim i As Integer, j As Integer
Dim INS() As Byte
    If MeMain = False Then Exit Sub
    If MSCommKQ16.CommPort = 16 Then Exit Sub
    
    Call Num_ALL        '统计
    Call NumNJX_ALL     '统计非机械报警点(包括报警和已经确认的报警)
    
    If MSCommKQ16.PortOpen = True And MeMain = True Then    '#Q0Nxxxx[0DH0AH]xx$
        '************************************************************************
        '在这里设置输出口
        '************************************************************************
        For i = 0 To 15
            Q16(i) = False
        Next i
        
        '蜂鸣器声报警
        Q16(0) = SONK
        '全船报警( 声 )( 蜂鸣后延迟1分钟全船声报警，可消声 )
        If timeSONK > 60 Then
            Q16(1) = True
        Else
            Q16(1) = False
        End If
        
        '台面综合报警指示灯(报警闪，消闪后平光)
        If EX_Alm > 0 Then
            Q16(2) = Not Q16(2)
        Else
            If EX_ACKAlm > 0 Then
                Q16(2) = True
            Else
                Q16(2) = False
            End If
        End If
        '延伸报警(有任一报警或已确认报警时动作,不可消)
        If EX_Alm + EX_ACKAlm <> 0 Then
            Q16(3) = True
        Else
            Q16(3) = False
        End If
        
        Q16(4) = SONKJX                 '机械声报警
        
        '***********************************************************
        '自定义报警点输出
        If MoniData(2, 2).Value < 1.6 Then Q16(5) = True    'ME LO INLET PRESS LOW
        If MoniData(2, 4).Value < 1.1 Then Q16(6) = True    'ME HT WATER JACKET INLET PRESS
        If MoniData(1, 14).Value > 92 Then Q16(7) = True    'ME HT WATER JACKET OUTLET TEMP
        If MoniData(2, 6).Value > 520 Or _
            MoniData(2, 7).Value > 520 Or _
            MoniData(2, 8).Value > 520 Or _
            MoniData(2, 9).Value > 520 Or _
            MoniData(2, 10).Value > 520 Or _
            MoniData(2, 11).Value > 520 Or _
            MoniData(2, 12).Value > 520 Or _
            MoniData(2, 13).Value > 520 Or _
            MoniData(2, 14).Value > 520 Then                'ME EXH GAS CYL 1-9 OUTLET TEMP
            Q16(10) = True
        End If
        '***********************************************************
            
        '组合输出字符串
        S1 = Bin2Asc(Not Q16(3), Not Q16(2), Not Q16(1), Not Q16(0))
        S2 = Bin2Asc(Not Q16(7), Not Q16(6), Not Q16(5), Not Q16(4))
        S3 = Bin2Asc(Not Q16(11), Not Q16(10), Not Q16(9), Not Q16(8))
        S4 = Bin2Asc(Not Q16(15), Not Q16(14), Not Q16(13), Not Q16(12))
        S = "#Q0N" & S1 & S2 & S3 & S4 & vbCrLf
        S = S & ASCSumF(S, 2) & "$"
        '呼叫 & 接收
        LabQ16 = Format(Now, "hh:mm:ss   ") & S
        MSCommKQ16.Output = S
    End If
    If MSCommKQ16.PortOpen = True Then
        S = MSCommKQ16.Input
        RTKQ16.SelStart = Len(RTKQ16.Text)
        RTKQ16.SelText = S          '文本显示
    End If
    
    KQCall.CallFail = KQCall.CallFail + 1
    '拆解字符串  #I0N??>?[ODHOAH]03$
    If TxtRead(RTKQ16, "#I", "$", 11, , True) = True Then  '处理数据
        INS = OutStr(0)
        For i = 0 To 3
            Call Byte2Bin(INS(6 + i * 2))
            For j = 0 To 3
                K16(i * 4 + j) = Not Bin(j)
            Next j
        Next i
        '显示输入值
        S = ""
        For i = 0 To 3
            For j = 0 To 3
                If K16(i * 4 + j) = True Then
                    S = S & "1"
                Else
                    S = S & "0"
                End If
            Next j
            S = S & ","
        Next i
        LabK16.Caption = Now & vbCrLf & S
        KQCall.CallFail = 0
    End If
    
    'KQ16 通讯故障判定
    If KQCall.CallFail >= KQCall.MaxCall Then
        KQCall.CallFail = 999
        If KQCall.CommFail = False Then
            KQCall.CommFail = True
            FormSelfCheck.Show
            SONK = True
        End If
        For i = 0 To 15
            K16(i) = False
        Next i
    Else
        KQCall.CommFail = False
    End If
    
    '************************************************************************
    '在这里设置输入处理程序
    '************************************************************************
    If K16(0) = True Then   '消声键
        SONK = False
        SONKJX = False
    End If
    
    If K16(1) = True Then   '消闪键
        Call ComACK_Click
    End If
    
'    If K16(2) = True Or CheckTestT(0).Value = 1 Then   'UPS故障报警
'        If UPSAlm = False Then  '上升沿触发
'            SONK = True
'            UPSAlm = True
'            FormSelfCheck.Show
'        End If
'    Else
'        UPSAlm = False
'    End If
'    If K16(3) = True Or CheckTestT(1).Value = 1 Then   '死人报警
'        If DManAlm = False Then '上升沿触发
'            SONK = True
'            DManAlm = True
'            FormSelfCheck.Show
'        End If
'    Else
'        DManAlm = False
'    End If
End Sub

Private Sub TimerMS_Timer()
Dim S As String
Dim MSA As Boolean, MSB As Boolean

    '判断主从 090706
'    If MSComm1.PortOpen = True Then
'        If MSComm1.CDHolding = False Then
'            If MeMain = False Then
'                MeMain = True
'                Call LoadSet     '如果是从机变为主机 就读入设置
'            End If
'            MeMain = True
'            MSComm1.DTREnable = True
'        Else
'            MeMain = False
'            MSComm1.DTREnable = False
'        End If
'    Else
'        MeMain = False
'    End If

    'commnuication fail??
    MS1Call.CallFail = MS1Call.CallFail + 1
    MS2Call.CallFail = MS2Call.CallFail + 1
    If MSComm1.PortOpen = True Then
        S = MSComm1.Input
        If S <> "" Then MS1Call.CallFail = 0
        MSComm1.Output = "*"
    End If
    If MSComm2.PortOpen = True Then
        S = MSComm2.Input
        If S <> "" Then MS2Call.CallFail = 0
        MSComm2.Output = "*"
    End If
    If MS1Call.CallFail > MS1Call.MaxCall Then
        MS1Call.CallFail = 999
        If MS1Call.CommFail = False Then
            SONK = True
            MS1Call.CommFail = True
            FormSelfCheck.Show
        End If
    Else
        MS1Call.CommFail = False
    End If
    If MS2Call.CallFail > MS2Call.MaxCall Then
        MS2Call.CallFail = 999
        If MS2Call.CommFail = False Then
            SONK = True
            MS2Call.CommFail = True
            FormSelfCheck.Show
        End If
    Else
        MS2Call.CommFail = False
    End If
    '判断主从
    If MSComm1.PortOpen = True Then
        If MSComm1.CDHolding = False Then
            MSA = True
        Else
            MSA = False
        End If
    End If
    If MSComm2.PortOpen = True Then
        If MSComm2.CDHolding = False Then
            MSB = True
        Else
            MSB = False
        End If
    End If
    '如果两个口都无法打开，我们认为程序处于手提电脑中，置为主站
    If MSComm1.PortOpen = False And MSComm2.PortOpen = False Then
        MSA = True: MSB = True
    End If
    If MSA = MSB Then
        If MeMain = False And MSA = True Then
            MeMain = True
            Call LoadSet     '如果是从机变为主机 就读入设置
        End If
        MeMain = MSA
    Else
        If MS1Call.CommFail = True And MS2Call.CommFail = False Then
            If MeMain = False And MSB = True Then
                MeMain = True
                Call LoadSet     '如果是从机变为主机 就读入设置
            End If
            MeMain = MSB
        End If
        If MS1Call.CommFail = False And MS2Call.CommFail = True Then
            If MeMain = False And MSA = True Then
                MeMain = True
                Call LoadSet     '如果是从机变为主机 就读入设置
            End If
            MeMain = MSA
        End If
    End If
    If MSComm1.PortOpen = True Then MSComm1.DTREnable = MeMain
    If MSComm2.PortOpen = True Then MSComm2.DTREnable = MeMain
End Sub

Private Sub TimerNow_Timer()
    LabDate.Caption = Format(Now, "yy-mm-dd")
    LabTime.Caption = Format(Now, "hh:mm:ss")
    
    If SONK = True Then     '计算声报警持续时间(S)
        timeSONK = timeSONK + 1
        If timeSONK > 999 Then timeSONK = 999
    Else
        timeSONK = 0
    End If
End Sub

Private Sub TimerQ32_Timer()
Dim S1 As String, S2 As String, S3 As String, S4 As String, S As String
Dim S5 As String, S6 As String, S7 As String, S8 As String
Dim N As Integer, i As Integer
Dim Q(31) As Boolean
    If MeMain = False Then Exit Sub
    If MSCommQ32.CommPort = 16 Then Exit Sub
    
    For i = 0 To 31
        Q(i) = False
    Next i
    '控制集控台指示灯
    Q(1) = BinData(2, 19).Alm Or BinData(2, 20).Alm Or _
            BinData(2, 21).Alm Or BinData(2, 22).Alm Or BinData(2, 23).Alm
    Q(2) = BinData(2, 15).Alm Or BinData(2, 28).Alm
    Q(3) = BinData(2, 0).Alm Or BinData(2, 1).Alm Or BinData(2, 2).Alm Or _
            MoniData(0, 12).Alm
    Q(4) = BinData(5, 22).Alm Or BinData(5, 23).Alm Or BinData(5, 24).Alm Or _
            MoniData(0, 13).Alm
    Q(5) = BinData(2, 16).Alm
    Q(6) = BinData(2, 29).Alm
    'Q(6) = BinData(2, 14).Alm Or BinData(2, 27).Alm Or BinData(8, 7).Alm Or BinData(8, 8).Alm Or _
               MoniData(0, 6).Alm Or BinData(0, 7).Alm Or MoniData(0, 10).Alm Or BinData(0, 11).Alm
    Q(7) = BinData(3, 0).Alm Or BinData(3, 1).Alm Or BinData(3, 2).Alm Or BinData(3, 3).Alm Or _
               BinData(3, 4).Alm Or BinData(3, 5).Alm Or BinData(3, 6).Alm Or BinData(3, 7).Alm Or _
                BinData(3, 8).Alm Or BinData(3, 9).Alm Or _
               MoniData(2, 8).Alm 'Or MoniData(0, 3).Alm
    Q(8) = BinData(5, 0).Alm Or BinData(5, 1).Alm Or BinData(5, 2).Alm Or BinData(5, 3).Alm Or _
               BinData(5, 26).Alm Or BinData(5, 5).Alm Or BinData(5, 6).Alm Or BinData(5, 7).Alm Or _
                BinData(5, 8).Alm Or BinData(5, 9).Alm Or _
               MoniData(3, 6).Alm 'Or MoniData(0, 19).Alm
   Q(9) = BinData(5, 11).Alm Or BinData(5, 12).Alm Or BinData(5, 13).Alm Or BinData(5, 14).Alm Or _
               BinData(5, 15).Alm Or BinData(5, 16).Alm Or BinData(5, 17).Alm Or BinData(5, 18).Alm Or _
                BinData(5, 19).Alm Or BinData(5, 20).Alm Or _
               MoniData(3, 7).Alm 'Or MoniData(0, 20).Alm
    Q(10) = BinData(9, 1).Alm Or BinData(9, 2).Alm Or BinData(9, 3).Alm Or BinData(9, 4).Alm Or BinData(9, 5).Alm Or BinData(9, 6).Alm Or BinData(9, 7).Alm Or BinData(9, 8).Alm Or BinData(9, 10).Alm
    Q(11) = BinData(1, 24).Alm Or BinData(1, 25).Alm Or BinData(1, 26).Alm
    Q(12) = BinData(1, 28).Alm Or BinData(1, 29).Alm Or BinData(1, 30).Alm
    Q(13) = BinData(6, 4).Alm Or BinData(6, 5).Alm Or BinData(4, 7).Alm
    Q(14) = BinData(3, 11).Alm Or BinData(3, 12).Alm Or BinData(3, 13).Alm Or BinData(3, 14).Alm Or _
             BinData(3, 15).Alm Or BinData(3, 16).Alm Or BinData(3, 17).Alm Or BinData(3, 18).Alm Or _
            BinData(3, 19).Alm Or BinData(3, 20).Alm Or BinData(3, 21).Alm Or BinData(3, 22).Alm Or _
            BinData(3, 23).Alm Or BinData(3, 24).Alm Or BinData(3, 25).Alm Or BinData(3, 26).Alm Or _
            BinData(3, 27).Alm Or BinData(3, 28).Alm Or BinData(3, 29).Alm Or BinData(3, 30).Alm Or _
            BinData(3, 31).Alm
    Q(15) = BinData(0, 0).Alm Or BinData(0, 1).Alm
    Q(22) = BinData(2, 25).Alm
    Q(23) = MoniData(0, 11).Alm
    Q(24) = MoniData(0, 13).Alm
    Q(25) = BinData(1, 22).Alm
    Q(26) = BinData(8, 8).Alm
    Q(27) = BinData(2, 28).Alm
    Q(16) = BinData(2, 12).Alm
    Q(17) = MoniData(0, 10).Alm
    Q(18) = MoniData(0, 12).Alm
    Q(19) = BinData(1, 23).Alm
    Q(20) = BinData(8, 7).Alm
    Q(21) = BinData(2, 28).Alm
    
    '组合输出字符串
    '#Q0N0?0?????
    'XX$
    S1 = Bin2Asc(Not Q(3), Not Q(2), Not Q(1), Not Q(0))
    S2 = Bin2Asc(Not Q(7), Not Q(6), Not Q(5), Not Q(4))
    S3 = Bin2Asc(Not Q(11), Not Q(10), Not Q(9), Not Q(8))
    S4 = Bin2Asc(Not Q(15), Not Q(14), Not Q(13), Not Q(12))
    S5 = Bin2Asc(Not Q(19), Not Q(18), Not Q(17), Not Q(16))
    S6 = Bin2Asc(Not Q(23), Not Q(22), Not Q(21), Not Q(20))
    S7 = Bin2Asc(Not Q(27), Not Q(26), Not Q(25), Not Q(24))
    S8 = Bin2Asc(Not Q(31), Not Q(30), Not Q(29), Not Q(28))
    S = "#Q" & N & "N" & S1 & S2 & S3 & S4 & S5 & S6 & S7 & S8 & vbCrLf
    S = S & ASCSumF(S, 2) & "$"
    
    If MSCommQ32.PortOpen = True Then
        MSCommQ32.Output = S
        LabQ32.Caption = Format(Now, "mm:ss ") & S
        
        S = MSCommQ32.Input
        If InStr(S, "Q0N") <> 0 Then Q32Call.CallFail = 0
    End If
    
    Q32Call.CallFail = Q32Call.CallFail + 1
    'Q32 通讯故障判定
    If Q32Call.CallFail >= Q32Call.MaxCall Then
        Q32Call.CallFail = 999
        If Q32Call.CommFail = False Then
            Q32Call.CommFail = True
            FormSelfCheck.Show
            SONK = True
        End If
    Else
        Q32Call.CommFail = False
    End If
End Sub

Private Sub TimerReset_Timer()
Dim i As Integer, j As Integer, N As Integer
Dim ALMTmp As Boolean
Dim SysAlm As Boolean
    ImageSONK.Visible = SONK
    LabDuty.Caption = WatchName(Watch1)
    
    Flash = Not Flash
    
    If MeMain = True Then
        ComSilence.Enabled = True
        ComACK.Enabled = True
        ComDuty.Enabled = True
        ComOption.Enabled = True
        ComCTime.Enabled = True
        ComSysAlm.Enabled = True
    Else
        ComSilence.Enabled = False
        ComACK.Enabled = False
        ComDuty.Enabled = False
        ComOption.Enabled = False
        ComCTime.Enabled = False
        ComSysAlm.Enabled = False
    End If
    
    If MeMain = True Then
        If Val(TimerReset.Tag) = 0 Then Call SendSlave  '如果为主机的话3秒写一次状态
        Call BinReset           'K32 set 刷新
        Call MoniReset          'AT24 set 刷新
        'Call PJReset            '平均温差
        If RP = True Then Call ListAlmReport
    Else
        Call ReadSlave                                  '如果为从机的话1秒从主机读一次状态
        Call BinResetB           'K32 set 刷新
        Call MoniResetB          'AT24 set 刷新
        'Call PJResetB           '平均温差 090706
        If RP = True Then Call ListAlmReport
    End If
    Call GroupReset         '页面(组)报警刷新
    
    '系统故障按钮
    SysAlm = False
    For N = BinCallSta To BinCallEnd
        If BinCall(N).CallFail >= BinCall(N).MaxCall Then SysAlm = True
    Next N
    For N = MoniCallSta To MoniCallEnd
        If MoniCall(N).CallFail >= MoniCall(N).MaxCall Then SysAlm = True
    Next N
    If MeMain = True Then
        For N = EMRCallSta To EMRCallEnd
            If EMRCall(N).CallFail >= EMRCall(N).MaxCall Then SysAlm = True
        Next N
    End If
    If MSCommKQ16.CommPort <> 16 And KQCall.CallFail >= KQCall.MaxCall Then SysAlm = True
    If MSCommQ32.CommPort <> 16 And Q32Call.CallFail >= Q32Call.MaxCall Then SysAlm = True
    If MSCommU24.CommPort <> 16 And U24Call.CallFail >= U24Call.MaxCall Then SysAlm = True
    If MS1Call.CommFail = True Or MS2Call.CommFail = True Then SysAlm = True
    If SysAlm = True And MeMain = True Then
        ComSysAlm.BackColor = vbRed
    Else
        ComSysAlm.BackColor = &H8000000F
    End If
    
    TimerReset.Tag = Val(TimerReset.Tag) + 1
    If Val(TimerReset.Tag) > 3 Then TimerReset.Tag = "0"
End Sub

Private Sub TimerSaveEnable_Timer()
    TimerSaveEnable.Enabled = False
End Sub

Private Sub TimerU24_Timer()
Dim i As Integer, j As Integer, x As Integer, y As Integer
Dim AD As Integer, ADU As Integer, ADL As Integer
Dim DA As Integer, DAU As Integer, DAL As Integer
Dim SendS As String, S As String, SRead As String
    If MeMain = False Then Exit Sub
    If MSCommU24.CommPort = 16 Then Exit Sub
    
    '计算DA
    For i = 0 To 0
        For j = 0 To 23
            If U24Data(i, j).DispImg = "9-99" Then
                U24Data(i, j).DA = 0
            Else
                x = Val(Left(U24Data(i, j).DispImg, 1))
                y = Val(Right(U24Data(i, j).DispImg, 2))
                AD = MoniData(x, y).AD
                ADU = MoniData(x, y).ADU
                ADL = MoniData(x, y).ADL
                DAU = U24Data(i, j).DAU
                DAL = U24Data(i, j).DAL
                DA = (AD - ADL) * ((DAU - DAL) / (ADU - ADL)) + DAL
                If DA > 256 Then DA = 256
                If DA < 0 Then DA = 0
                U24Data(i, j).DA = DA
            End If
        Next j
    Next i
    '组合发送字符串
    SendS = "#Q0M "
    i = 0
    For j = 0 To 23
        S = Hex(U24Data(i, j).DA)
        S = Right("00" & S, 2)
        Select Case Left(S, 1)
            Case "A": SendS = SendS & ":"
            Case "B": SendS = SendS & ";"
            Case "C": SendS = SendS & "<"
            Case "D": SendS = SendS & "="
            Case "E": SendS = SendS & ">"
            Case "F": SendS = SendS & "?"
            Case Else: SendS = SendS & Left(S, 1)
        End Select
        Select Case Right(S, 1)
            Case "A": SendS = SendS & ":"
            Case "B": SendS = SendS & ";"
            Case "C": SendS = SendS & "<"
            Case "D": SendS = SendS & "="
            Case "E": SendS = SendS & ">"
            Case "F": SendS = SendS & "?"
            Case Else: SendS = SendS & Right(S, 1)
        End Select
    Next j
    SendS = SendS & vbCrLf
    SendS = SendS & ASCSumF(SendS, 2) & "$"
    
    'S = "#Q0M 010203040506070809101112131415161718192021222324" & vbCrLf
    If MSCommU24.PortOpen = True And MeMain = True Then
        MSCommU24.Output = SendS
        LabU24.Caption = Format(Now, "hh:mm:ss  ") & SendS
    End If
    
    U24Call.CallFail = U24Call.CallFail + 1
    If MSCommU24.PortOpen = True Then
        SRead = MSCommU24.Input
        If Left(SRead, 4) = "*Q0M" Then U24Call.CallFail = 0
    End If
    
    'U24 通讯故障判定
    If U24Call.CallFail >= U24Call.MaxCall Then
        U24Call.CallFail = 999
        If U24Call.CommFail = False Then
            U24Call.CommFail = True
            FormSelfCheck.Show
            SONK = True
        End If
    Else
        U24Call.CommFail = False
    End If
End Sub

Private Sub TimerVDR_Timer()
Dim S As String, SM As String, Stmp As String
Dim i As Integer
Dim x As Single
    If MeMain = False Then Exit Sub
    If MSCommVDR.CommPort = 16 Then Exit Sub
    
    i = 0
    Do While VDRAddB(i) >= 0
        If BinData(VDRAddB(i) \ 32, VDRAddB(i) Mod 32).SF = True Then
            S = S & "-"
        Else
            If BinData(VDRAddB(i) \ 32, VDRAddB(i) Mod 32).Alm = True Then
                S = S & "1"
            Else
                S = S & "0"
            End If
        End If
        i = i + 1
    Loop
    S = "(BIN:" & S & ")"
    
    i = 0
    Do While VDRAddM(i) >= 0
        If MoniData(VDRAddM(i) \ 24, VDRAddM(i) Mod 24).SF = True Then
            Stmp = "------"
        Else
            x = MoniData(VDRAddM(i) \ 24, VDRAddM(i) Mod 24).Value
            If x >= 1000 Then Stmp = Format(x, "0000.0")
            If x < 1000 And x >= 0 Then Stmp = Format(x, "000.00")
            If x < 0 Then Stmp = Format(x, "00.00")
        End If
        
        SM = SM & Stmp & ","
        i = i + 1
    Loop
    SM = "[MONI:" & SM & "]"
    
    If MSCommVDR.PortOpen = True Then
        MSCommVDR.Output = S & vbCrLf & SM
        LabVDR(0) = Format(Now, "hh:mm:ss  ") & S
        LabVDR(1) = Format(Now, "hh:mm:ss  ") & SM
    End If
End Sub

Public Sub GroupReset()             '页面(组)报警刷新
Dim i As Integer, j As Integer
Dim AX As Integer, AY As Integer    '地址
Dim PN As Integer                   '控件标号
Dim S As String, SL As String
    For i = 1 To 7
        PageAlm(i) = False          '默认为无报警
        PageExAlm(i) = False
        For j = 0 To ListNum - 1
            PN = (i - 1) * ListNum + j
            AX = PxAL(i, j).AddX
            AY = PxAL(i, j).AddY
            Select Case PxAL(i, j).BorM
            Case "B"
                ShapePAlm(PN).Visible = True
                LabPList(PN).Visible = True
                '刷新文字和指示灯
                If LabPList(PN).Caption <> BinData(AX, AY).Name Then LabPList(PN).Caption = BinData(AX, AY).Name
                ShapePAlm(PN).FillColor = ShapeAlm(AX * 32 + AY).FillColor
                '页面报警判定
                SL = LabCHName(AX * 32 + AY).ToolTipText
                If SL = "ALM" Or SL = "SF" Then PageAlm(i) = True
                If SL = "ALM" Or SL = "ACKN-ALM" Then PageExAlm(i) = True
            Case "M"
                ShapePAlm(PN).Visible = True
                LabPList(PN).Visible = True
                '页面(组)报警判定
                SL = LabCHNameM(AX * 24 + AY).ToolTipText
                If SL = "ALM" Or SL = "SF" Then PageAlm(i) = True
                If SL = "ALM" Or SL = "ACKN-ALM" Then PageExAlm(i) = True
                '刷新文字和指示灯
                If MoniData(AX, AY).SF = False And SL <> "SF" And SL <> "ACKN-SF" Then
                    S = MoniData(AX, AY).Name & " " & Format(MoniData(AX, AY).Value, MoniData(AX, AY).FmtS) & MoniData(AX, AY).Unit
                Else
                    S = MoniData(AX, AY).Name & " S.F."
                End If
                If MoniData(AX, AY).UseBin = True Then S = MoniData(AX, AY).Name    '2009-01-08
                
                If LabPList(PN).Caption <> S Then LabPList(PN).Caption = S
                ShapePAlm(PN).FillColor = ShapeAlmM(AX * 24 + AY).FillColor
            Case Else
                ShapePAlm(PN).Visible = False
                LabPList(PN).Visible = False
            End Select
        Next j
        If PageAlm(i) = True Then   '闪烁
            If Flash = True Then
                ComGroup(i).BackColor = &H8000000F
            Else
                ComGroup(i).BackColor = vbRed
            End If
        Else
            If PageExAlm(i) = True Then
                ComGroup(i).BackColor = vbRed
            Else
                ComGroup(i).BackColor = &H8000000F
            End If
        End If
    Next i

End Sub

Public Sub BinReset()           '开关量界面刷新
Dim i As Integer, N As Integer
Dim NeedSave As Boolean
Dim Prn As String
Dim S As String
Dim x As Integer, y As Integer, Ct As Boolean
On Error Resume Next
With Me
    For N = 0 To 9
    For i = 0 To 31
        NeedSave = False
        If CheckName.Value = 1 Then
            S = N & "-" & Format(i, "00")
        Else
            S = Left(BinData(N, i).Name, 6)
        End If
        If .LabCHName(N * 32 + i).Caption <> S Then .LabCHName(N * 32 + i).Caption = S
        
        Ct = False
        If BinData(N, i).CutImg <> "9-99" Then
            x = Val(Left(BinData(N, i).CutImg, 1))
            y = Val(Right(BinData(N, i).CutImg, 2))
            If BinData(x, y).Alm = False And BinData(x, y).SF = False Then
                Ct = True
            End If
        End If
        If BinData(N, i).Cutout = True Or Ct = True Then
            .LabCHName(N * 32 + i).ToolTipText = "Cutout"   '越控
        Else
            If .LabCHName(N * 32 + i).ToolTipText = "Cutout" Then   '解除越控
                .LabCHName(N * 32 + i).ToolTipText = "NR"
            End If
        End If
        Select Case .LabCHName(N * 32 + i).ToolTipText
            Case "Cutout"
                .ShapeAlm(N * 32 + i).FillColor = &HC0C0C0
            Case "NR", ""
                .ShapeAlm(N * 32 + i).FillColor = vbWhite
                If BinData(N, i).Alm = True Then
                    If BinData(N, i).Group <> 9 Then
                        .LabCHName(N * 32 + i).ToolTipText = "ALM"
                        BinData(N, i).AlmTime = Format(Now, tFmsL)
                        SONK = True: NeedSave = True: RP = True
                        Call AlmListAdd("B", N, i, "F")
                    Else    '运行指示 组9
                        .LabCHName(N * 32 + i).ToolTipText = "RUN"
                    End If
                    If IsJX(N, i) = True Then SONKJX = True
                End If
                If BinData(N, i).SF = True Then
                    .LabCHName(N * 32 + i).ToolTipText = "SF"
                    BinData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If IsJX(N, i) = True Then SONKJX = True
                End If
            Case "RUN"
                .ShapeAlm(N * 32 + i).FillColor = vbGreen
                If BinData(N, i).Alm = False And BinData(N, i).SF = False Then
                    .LabCHName(N * 32 + i).ToolTipText = "NR"
                End If
                If BinData(N, i).SF = True Then
                    .LabCHName(N * 32 + i).ToolTipText = "SF"
                    BinData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If IsJX(N, i) = True Then SONKJX = True
                End If
            Case "ALM"      '闪烁
                If Flash = True Then
                    .ShapeAlm(N * 32 + i).FillColor = vbRed
                Else
                    .ShapeAlm(N * 32 + i).FillColor = vbButtonFace
                End If
                If BinData(N, i).SF = True Then
                    .LabCHName(N * 32 + i).ToolTipText = "SF"
                    BinData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If IsJX(N, i) = True Then SONKJX = True
                End If
            Case "SF"       '闪烁
                If Flash = True Then
                    .ShapeAlm(N * 32 + i).FillColor = vbYellow
                Else
                    .ShapeAlm(N * 32 + i).FillColor = vbButtonFace
                End If
            Case "ACKN-ALM"
                .ShapeAlm(N * 32 + i).FillColor = vbRed
                If BinData(N, i).Alm = False And BinData(N, i).SF = False Then
                    .LabCHName(N * 32 + i).ToolTipText = "NR"
                    .ShapeAlm(N * 32 + i).FillColor = vbWhite
                    If BinData(N, i).Group <> 9 Then NeedSave = True
                End If
                If BinData(N, i).SF = True Then
                    .LabCHName(N * 32 + i).ToolTipText = "SF"
                    BinData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If IsJX(N, i) = True Then SONKJX = True
                End If
            Case "ACKN-SF"
                .ShapeAlm(N * 32 + i).FillColor = vbYellow
                If BinData(N, i).SF = False Then
                    .LabCHName(N * 32 + i).ToolTipText = "NR"
                    .ShapeAlm(N * 32 + i).FillColor = vbWhite
                    If BinData(N, i).Group <> 9 Then NeedSave = True
                End If
        End Select
        If NeedSave = True And TimerSaveEnable.Enabled = False And MeMain = True Then
            '...............存入数据库
            FormRecord.Adodc1.Recordset.AddNew
            FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
            Prn = Format(Now, tFmsL)
            FormRecord.Adodc1.Recordset("Name") = Left(BinData(N, i).Name, 50)
            Prn = Prn & vbTab & BinData(N, i).Name
            If BinData(N, i).SF = True Then
                FormRecord.Adodc1.Recordset("Value") = "Sensor Fail"
                Prn = Prn & vbTab & "Sensor Fail"
            Else
                If BinData(N, i).Alm = True Then
                    FormRecord.Adodc1.Recordset("Value") = "ALM"
                    Prn = Prn & vbTab & "ALM"
                Else
                    FormRecord.Adodc1.Recordset("Value") = "NOR"
                    Prn = Prn & vbTab & "NOR"
                End If
            End If
            FormRecord.Adodc1.Recordset.Update
            DoEvents
            Open "prn" For Output As #1
            If MeMain = True Then
                If AutoPrint = True Then Print #1, Prn
                ListSlaveSave.AddItem Prn
            End If
            Close #1
            If AutoZorder = True Then FormRecord.Show
        End If
    Next i
    Next N
End With
End Sub

Public Sub MoniReset()          '模拟量界面刷新
Dim i As Integer, N As Integer
Dim S As String, S1 As String, S2 As String
Dim NeedSave As Boolean
Dim Prn As String
Dim x As Integer, y As Integer, Ct As Boolean
On Error Resume Next
With Me
    For N = 0 To 9
    For i = 0 To 23
        NeedSave = False
        If .CheckAD.Value = 1 Then
            S1 = Format(MoniData(N, i).AD, "0000")
        Else
            S1 = Format(MoniData(N, i).Value, MoniData(N, i).FmtS) & MoniData(N, i).Unit
        End If
        If .CheckName.Value = 1 Then
            S2 = N & "-" & Format(i, "00")
        Else
            S2 = MoniData(N, i).Name
        End If
        S = Left(S1 & " " & S2, 12)
        If .LabCHNameM(N * 24 + i).Caption <> S Then .LabCHNameM(N * 24 + i).Caption = S
        
        Ct = False
        If MoniData(N, i).CutImg <> "9-99" Then
            x = Val(Left(MoniData(N, i).CutImg, 1))
            y = Val(Right(MoniData(N, i).CutImg, 2))
            If BinData(x, y).Alm = False And BinData(x, y).SF = False Then
                Ct = True
            End If
        End If
        If MoniData(N, i).Cutout = True Or Ct = True Then
            .LabCHNameM(N * 24 + i).ToolTipText = "Cutout"
            .ShapeAlmM(N * 24 + i).FillColor = &HC0C0C0
        Else
        Select Case .LabCHNameM(N * 24 + i).ToolTipText
            Case "Cutout"
                .LabCHNameM(N * 24 + i).ToolTipText = "NR"
                .ShapeAlmM(N * 24 + i).FillColor = &HC0C0C0
            Case "NR", ""
                If MoniData(N, i).Alm = True Then
                    .LabCHNameM(N * 24 + i).ToolTipText = "ALM"
                    MoniData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    Call AlmListAdd("M", N, i, "F")
                    If N * 24 + i <> 77 Then SONKJX = True
                End If
                If MoniData(N, i).SF = True Then
                    .LabCHNameM(N * 24 + i).ToolTipText = "SF"
                    MoniData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If N * 24 + i <> 77 Then SONKJX = True
                End If
                If MoniData(N, i).Alm = False And MoniData(N, i).SF = False Then
                    .ShapeAlmM(N * 24 + i).FillColor = vbWhite
                End If
            Case "ALM"      '闪烁
                If MoniData(N, i).SF = True Then
                    .LabCHNameM(N * 24 + i).ToolTipText = "SF"
                    MoniData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If N * 24 + i <> 77 Then SONKJX = True
                End If
                If Flash = True Then
                    .ShapeAlmM(N * 24 + i).FillColor = vbRed
                Else
                    .ShapeAlmM(N * 24 + i).FillColor = vbButtonFace
                End If
            Case "SF"       '闪烁
                If Flash = True Then
                    .ShapeAlmM(N * 24 + i).FillColor = vbYellow
                Else
                    .ShapeAlmM(N * 24 + i).FillColor = vbButtonFace
                End If
            Case "ACKN-ALM"
                If MoniData(N, i).SF = True Then
                    .LabCHNameM(N * 24 + i).ToolTipText = "SF"
                    MoniData(N, i).AlmTime = Format(Now, tFmsL)
                    SONK = True: NeedSave = True: RP = True
                    If N * 24 + i <> 77 Then SONKJX = True
                End If
                If MoniData(N, i).Alm = False Then
                    .LabCHNameM(N * 24 + i).ToolTipText = "NR"
                    .ShapeAlmM(N * 24 + i).FillColor = vbWhite
                    NeedSave = True
                End If
            Case "ACKN-SF"
                If MoniData(N, i).SF = False Then
                    .LabCHNameM(N * 24 + i).ToolTipText = "NR"
                    .ShapeAlmM(N * 24 + i).FillColor = vbWhite
                    NeedSave = True
                End If
        End Select
        End If
        If NeedSave = True And TimerSaveEnable.Enabled = False And MeMain = True Then
            '...............存入数据库
            FormRecord.Adodc1.Recordset.AddNew
            FormRecord.Adodc1.Recordset("Time") = Format(Now, tFmsL)
            Prn = Format(Now, tFmsL)
            FormRecord.Adodc1.Recordset("Name") = Left(MoniData(N, i).Name, 50)
            Prn = Prn & vbTab & MoniData(N, i).Name
            If MoniData(N, i).SF = True Then
                FormRecord.Adodc1.Recordset("Value") = "Sensor Fail"
                Prn = Prn & vbTab & "Sensor Fail"
            Else
                FormRecord.Adodc1.Recordset("Value") = "ALM"
                Prn = Prn & vbTab & "ALM"
            End If
            FormRecord.Adodc1.Recordset.Update
            DoEvents
            Open "prn" For Output As #1
            If MeMain = True Then
                If AutoPrint = True Then Print #1, Prn
                ListSlaveSave.AddItem Prn
            End If
            Close #1
            If AutoZorder = True Then FormRecord.Show
        End If
    Next i
    Next N
End With
End Sub

Public Sub BinResetB()           '开关量界面刷新(只刷颜色，不改状态)
Dim i As Integer, N As Integer
Dim S As String
With Me
    For N = 0 To 9
    For i = 0 To 31
        If CheckName.Value = 1 Then
            S = N & "-" & Format(i, "00")
        Else
            S = Left(BinData(N, i).Name, 6)
        End If
        If .LabCHName(N * 32 + i).Caption <> S Then .LabCHName(N * 32 + i).Caption = S
        
        Select Case .LabCHName(N * 32 + i).ToolTipText
            Case "Cutout"
                .ShapeAlm(N * 32 + i).FillColor = &HC0C0C0
            Case "NR", ""
                .ShapeAlm(N * 32 + i).FillColor = vbWhite
            Case "RUN"
                .ShapeAlm(N * 32 + i).FillColor = vbGreen
            Case "ALM"      '闪烁
                If Flash = True Then
                    .ShapeAlm(N * 32 + i).FillColor = vbRed
                Else
                    .ShapeAlm(N * 32 + i).FillColor = vbButtonFace
                End If
            Case "SF"       '闪烁
                If Flash = True Then
                    .ShapeAlm(N * 32 + i).FillColor = vbYellow
                Else
                    .ShapeAlm(N * 32 + i).FillColor = vbButtonFace
                End If
            Case "ACKN-ALM"
                .ShapeAlm(N * 32 + i).FillColor = vbRed
            Case "ACKN-SF"
                .ShapeAlm(N * 32 + i).FillColor = vbYellow
        End Select
    Next i
    Next N
End With
End Sub

Public Sub MoniResetB()          '模拟量界面刷新(只刷颜色，不改状态)
Dim i As Integer, N As Integer
Dim NeedSave As Boolean
Dim S As String, S1 As String, S2 As String
With Me
    For N = 0 To 9
    For i = 0 To 23
        If .CheckAD.Value = 1 Then
            S1 = Format(MoniData(N, i).AD, "0000")
        Else
            S1 = Format(MoniData(N, i).Value, MoniData(N, i).FmtS) & MoniData(N, i).Unit
        End If
        If .CheckName.Value = 1 Then
            S2 = N & "-" & Format(i, "00")
        Else
            S2 = MoniData(N, i).Name
        End If
        S = Left(S1 & " " & S2, 12)
        If .LabCHNameM(N * 24 + i).Caption <> S Then .LabCHNameM(N * 24 + i).Caption = S
        
        Select Case .LabCHNameM(N * 24 + i).ToolTipText
            Case "Cutout"
                .ShapeAlmM(N * 24 + i).FillColor = &HC0C0C0
            Case "NR", ""
                .ShapeAlmM(N * 24 + i).FillColor = vbWhite
            Case "ALM"      '闪烁
                If Flash = True Then
                    .ShapeAlmM(N * 24 + i).FillColor = vbRed
                Else
                    .ShapeAlmM(N * 24 + i).FillColor = vbButtonFace
                End If
            Case "SF"       '闪烁
                If Flash = True Then
                    .ShapeAlmM(N * 24 + i).FillColor = vbYellow
                Else
                    .ShapeAlmM(N * 24 + i).FillColor = vbButtonFace
                End If
            Case "ACKN-ALM"
                .ShapeAlmM(N * 24 + i).FillColor = vbRed
            Case "ACKN-SF"
                .ShapeAlmM(N * 24 + i).FillColor = vbYellow
        End Select
    Next i
    Next N
End With
End Sub

'写状态副本送给从机
Public Sub SendSlave()
Dim i As Integer, j As Integer
Dim S As String
    ListSlave.Clear
    For i = 0 To 9
        For j = 0 To 31
            S = i & vbTab & j & vbTab & LabCHName(i * 32 + j).ToolTipText
            ListSlave.AddItem S
        Next j
    Next i
    For i = 0 To 9
        For j = 0 To 23
            S = i & vbTab & j & vbTab & LabCHNameM(i * 24 + j).ToolTipText
            S = S & vbTab & MoniData(i, j).Value
            ListSlave.AddItem S
        Next j
    Next i
    S = LabTYP(0).ToolTipText   '平均温差  20090625
    For i = 1 To 8
        S = S & vbTab & LabTYP(i).ToolTipText
    Next i
    ListSlave.AddItem S
    '值班人
    S = Watch1 & vbTab & Watch2 & vbTab & DealyW1 & vbTab & DealyW2 & vbTab & DealyW3 & vbTab & Format(Now, "yy-mm-dd hh:mm:ss")
    ListSlave.AddItem S
    Call List2File(ListSlave, App.Path & "\SDisp.ini")
    
    Call List2File(ListSlaveSave, App.Path & "\SSave.ini")
    ListSlaveSave.Clear
End Sub

'读状态副本从主机
Public Sub ReadSlave()
Dim PathMaMa As String
Dim FileMaMa As String, FileBaby As String
Dim k As Integer, i As Integer, j As Integer
Dim MaxDy As Integer
On Error GoTo getfail
    If CheckMain.Value = 1 Then
        PathMaMa = IntDirSlave
    Else
        PathMaMa = IntDirMain
    End If
    FileMaMa = PathMaMa & "\SDisp.ini": FileBaby = App.Path & "\GDisp.ini": FileCopy FileMaMa, FileBaby: Kill FileMaMa
    FileMaMa = PathMaMa & "\SSave.ini": FileBaby = App.Path & "\GSave.ini": FileCopy FileMaMa, FileBaby: Kill FileMaMa
    '如果成功读取次数超过60次则认为主从状态稳定，开始同步设置值
    GetWinTime = GetWinTime + 1
    If GetWinTime > 60 Then
        GetWinTime = 60
        FileMaMa = PathMaMa & "\SetB.ini": FileBaby = App.Path & "\SetB.ini": FileCopy FileMaMa, FileBaby
        FileMaMa = PathMaMa & "\SetM.ini": FileBaby = App.Path & "\SetM.ini": FileCopy FileMaMa, FileBaby
        'FileMaMa = PathMaMa & "\SetAPL.ini": FileBaby = App.Path & "\SetAPL.ini": FileCopy FileMaMa, FileBaby
        'FileMaMa = PathMaMa & "\SetSys.ini": FileBaby = App.Path & "\SetSys.ini": FileCopy FileMaMa, FileBaby
        'FileMaMa = PathMaMa & "\SetU.ini": FileBaby = App.Path & "\SetU.ini": FileCopy FileMaMa, FileBaby
        'FileMaMa = PathMaMa & "\VDRB.ini": FileBaby = App.Path & "\VDRB.ini": FileCopy FileMaMa, FileBaby
        'FileMaMa = PathMaMa & "\VDRM.ini": FileBaby = App.Path & "\VDRM.ini": FileCopy FileMaMa, FileBaby
    End If
    GetFTime = 0
    FrameW.Visible = False
    
    If GetWinTime < 5 Then   '对5次以内读到的文件不处理，以免时间同步错误
        Exit Sub
    End If
    
    Call File2List(App.Path & "\GDisp.ini", FormMain.ListTemp)
    For k = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = k
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) = 2 Then
            i = FanWei(Val(OutStr(0)), 0, 19)
            j = FanWei(Val(OutStr(1)), 0, 31)
            LabCHName(i * 32 + j).ToolTipText = OutStr(2)
            Select Case LabCHName(i * 32 + j).ToolTipText
            Case "Cutout"
            Case "NR", ""
                BinData(i, j).Alm = False: BinData(i, j).SF = False
            Case "RUN", "ALM", "ACKN-ALM"
                BinData(i, j).Alm = True: BinData(i, j).SF = False
            Case "SF", "ACKN-SF"
                BinData(i, j).SF = True
            End Select
        End If
        If UBound(OutStr) = 3 Then
            i = FanWei(Val(OutStr(0)), 0, 19)
            j = FanWei(Val(OutStr(1)), 0, 23)
            LabCHNameM(i * 24 + j).ToolTipText = OutStr(2)
            Select Case LabCHNameM(i * 24 + j).ToolTipText
            Case "Cutout"
            Case "NR", ""
                MoniData(i, j).Alm = False: MoniData(i, j).SF = False
            Case "RUN", "ALM", "ACKN-ALM"
                MoniData(i, j).Alm = True: MoniData(i, j).SF = False
            Case "SF", "ACKN-SF"
                MoniData(i, j).SF = True
            End Select
            MoniData(i, j).Value = OutStr(3)
        End If
        If UBound(OutStr) = 5 Then
            Watch1 = FanWei(Val(OutStr(0)), 0, 4)
            Watch2 = FanWei(Val(OutStr(1)), 0, 4)
            DealyW1 = FanWei(Val(OutStr(2)), 0, 4)
            DealyW2 = FanWei(Val(OutStr(3)), 0, 4)
            DealyW3 = FanWei(Val(OutStr(4)), 0, 4)
            SaveSetting "RDMS System", "Duty", "Watch1", Watch1
            SaveSetting "RDMS System", "Duty", "Watch2", Watch2
            SaveSetting "RDMS System", "Duty", "DealyW1", DealyW1
            SaveSetting "RDMS System", "Duty", "DealyW2", DealyW2
            SaveSetting "RDMS System", "Duty", "DealyW3", DealyW3
            Date = Left(OutStr(5), 8)
            Time = Right(OutStr(5), 8)
        End If
        If UBound(OutStr) = 8 Then      '20090625
            For i = 0 To 8
                LabTYP(i).ToolTipText = OutStr(i)
            Next i
        End If
    Next k
    
    Call File2List(App.Path & "\GSave.ini", FormMain.ListTemp)
    For k = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = k
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) = 2 Then
            FormRecord.Adodc1.Recordset.AddNew
            FormRecord.Adodc1.Recordset("Time") = OutStr(0)
            FormRecord.Adodc1.Recordset("Name") = Left(OutStr(1), 50)
            FormRecord.Adodc1.Recordset("Value") = OutStr(2)
            FormRecord.Adodc1.Recordset.Update
        End If
    Next k
    
    Exit Sub
    
getfail:
    GetFTime = GetFTime + 1
    If GetFTime > 999 Then GetFTime = 999
    
    If GetFTime > 10 Then
        If GetFTime > 30 Then
            LabW.Caption = "Loading data from Network fail.!!!(" & PathMaMa & ")"
            FrameW.Visible = True
            FrameW.ZOrder 0
        Else
            LabW.Caption = "Loading data from Network......(" & PathMaMa & ")(D:" & GetFTime & ")"
            FrameW.Visible = True
            FrameW.ZOrder 0
        End If
    End If
    '090706
    'If CheckMain.Value = 1 Then
    '    MaxDy = 20
    'Else
    '    MaxDy = 40
    'End If
    'If GetFTime >= MaxDy Then
    '    MeMain = True
    '    ComSilence.Enabled = True
    '    ComACK.Enabled = True
    '    ComDuty.Enabled = True
    '    ComOption.Enabled = True
    '    ComCTime.Enabled = True
    '    ComSysAlm.Enabled = True
    '
    '    GetFTime = 999
    '    GetWinTime = 0
    '    Call LoadSet
    '    Unload FormWait
    'End If
End Sub
