Attribute VB_Name = "Public"
Option Explicit

Public Const Password As String = "1234"    'ͨ������
Public Const Maxdelay As Integer = 29       '����ӳ��������
Public Const ListNum As Integer = 100       '����ҳ�浥ҳ��������
Public Const tFmsL As String = "dd-mm-yyyy hh:mm:ss"    'ʱ���ʽ
Public Const tFms1 As String = "dd-mm-yyyy"             'ʱ���ʽ
Public Const tFms2 As String = "mm-yyyy"                'ʱ���ʽ

Public IntDirMain As String                 'Ĭ�ϵ���PC����������ַ
Public IntDirSlave As String                'Ĭ�ϵĴ�PC����������ַ

Public MeMain As Boolean            '������־
Public GetFTime As Integer          '���ļ�ʧ�ܴ����������жϣ�
Public GetWinTime As Integer        '���ļ��ɹ������������жϣ�

Public AutoPrint As Boolean         '�Զ���ӡ
Public AutoZorder As Boolean        '�Զ���ʾ������¼
Public Flash As Boolean             '��˸��־λ

Public PageAlm(1 To 7) As Boolean        'ҳ�汨����־
Public PageExAlm(1 To 7) As Boolean
Public PageNow As Integer
Public PxAL(1 To 7, 0 To ListNum - 1) As PxAlmList  'ÿҳ100��
Type PxAlmList
    BorM As String      'B = ������  M = ģ����  �����ַ� = ����ʾ
    AddX As Integer     '��ַ
    AddY As Integer
    t As String         '����ʱ��
End Type

Public BinData(9, 31) As BinType    '���������� 0-9��̨ 0-31ͨ��
Type BinType
    Name As String              '��������
    Alm As Boolean              '����̬
    Group As Integer            '������ 0����,1-8��,9����
    Delay As Integer            '�ӳ�������� 0-29
    DelayA(Maxdelay) As Boolean '�ӳ������ջ FIFO
    Nor As Boolean              '��ת���
    SF As Boolean               '�Ƿ����
    Cutout As Boolean           '����
    CutImg  As String           '����ӳ��
    AlmTime As String           '���һ�α���ʱ��
End Type

Public MoniData(9, 23) As MoniType  'ģ�������� 0-9��̨ 0-23ͨ��
Type MoniType
    AD As Integer               'ADֵ         0000-4095
    ADU As Integer              '������ADֵ
    ADC As Integer              '�г�ADֵ
    ADL As Integer              '��λADֵ
    SFU As Integer              '��λ����ADֵ
    SFL As Integer              '��λ����ADֵ
    Cutout As Boolean           '����
    CutImg  As String           '����ӳ��
    DispU As Single             '��������ʾֵ
    DispL As Single             '��λ��ʾֵ
    Unit As String              '��λ
    FmtS As String              '��ʽ���ַ���
    
    Value As Single             '����ֵ
    Delay As Integer            '�ӳ�������� 0-29
    DelayV(Maxdelay) As Single  '�ӳ������ջ FIFO
    AlmU As Single              '�߱���ֵ
    AlmL As Single              '�ͱ���ֵ
    Alm As Boolean              '����̬
    Group As Integer            '������ 0����,1-8��,9����
    SF As Boolean               '�Ƿ����
    
    UseBin As Boolean           '�Ƿ���Ϊ��������ʹ�� '2009-01-08
    Name As String              '��������
    AlmTime As String           '���һ�α���ʱ��
End Type

Public U24Data(0, 23) As U24Type    'ģ�������� 0-1��̨ 0-23ͨ��
Type U24Type
    DA As Integer           'DAֵ         000-255
    DAU As Integer          '20mA DAֵ
    DAL As Integer          '4mA  DAֵ
    DispImg As String       '���ģ��������
End Type

Public PageGroup(1 To 8) As PageGroupType   '��ʾ������� 1-8 ҳ
Type PageGroupType
    Name As String          '��������
    Disp As Boolean         '�����Ƿ���ʾ
    ExtALm As Boolean       '���챨��̬(�б�����λ/ȫ��������)
    NewAlm As Boolean       '�±�����λ/���ͺ�����
End Type

Public PJ(8) As PJType      '����ƽ���²�
Type PJType
    Value As Single         'ƫ��ֵ
    DispT As Single         '��ʾֵ
    Alm As Boolean          '����̬
    SF As Boolean           '����̬
End Type
Public PJTemp As Single     '����ƽ���¶�
Public PJAlm As Single      '����ֵ

'ͨѶ״̬
Public BinCall(9) As CommCall   '������
Public MoniCall(9) As CommCall  'ģ����
Public EMRCall(15) As CommCall  '�����
Public KQCall As CommCall       'KQ16���
Public Q32Call As CommCall
Public U24Call As CommCall      'U24���
Public MS1Call As CommCall      'Main / Slave
Public MS2Call As CommCall

Type CommCall
    CallFail As Integer             '����ʧ�ܴ���
    MaxCall As Integer              '�����д���
    CommFail As Boolean             'ͨѶ����
End Type

Public BinCallSta As Integer        '��ʼ̨��
Public BinCallEnd As Integer        'ĩβ̨��
Public BinCallNo As Integer         '��������ǰ���к�

Public MoniCallSta As Integer       '��ʼ̨��
Public MoniCallEnd As Integer       'ĩβ̨��
Public MoniCallNo As Integer        'ģ������ǰ���к�

Public EMRCallSta As Integer        '��ʼ̨��
Public EMRCallEnd As Integer        'ĩβ̨��
Public EMRCallNo As Integer         '����嵱ǰ���к�

Public EMRNo(9) As Integer          '���������е�ַ
Public EMRTmN1 As Integer           '��ʱ��Сѭ������
Public EMRTmN2 As Integer           '��ʱ����ѭ������(������ֵһ��)

Public SysName As String            'ϵͳ����
Public SONK As Boolean              '������
Public RP As Boolean                '�Ƿ񴥷�FormReport

Public K16(15) As Boolean   'KQ16�����
Public Q16(15) As Boolean   'KQ16�����
Public UPSAlm As Boolean    'UPS��Դ
Public DManAlm As Boolean   '���˱���
Public timeSONK As Single   '���г���ʱ��
Public SONKJX As Boolean    '��е������

Public Sub LoadSet()
Dim i As Integer
Dim S As String
    With FormMain
    If .MSCommK32.PortOpen = True Then .MSCommK32.PortOpen = False
    If .MSCommAT24.PortOpen = True Then .MSCommAT24.PortOpen = False
    If .MSCommKQ16.PortOpen = True Then .MSCommKQ16.PortOpen = False
    If .MSCommQ32.PortOpen = True Then .MSCommQ32.PortOpen = False
    If .MSCommEMR.PortOpen = True Then .MSCommEMR.PortOpen = False
    If .MSCommU24.PortOpen = True Then .MSCommU24.PortOpen = False
    If .MSCommVDR.PortOpen = True Then .MSCommVDR.PortOpen = False
    End With
    
    Call SetGroup       '��������ʼ��
    Call ReadSys        '��ȡϵͳ����
    Call ReadBin        '��ȡ����������
    Call ReadMoni       '��ȡģ��������
    MoniData(3, 18).UseBin = True       '3-18Ϊ���� '2009-01-08
    MoniData(3, 19).UseBin = True       '3-19Ϊ���� '2009-01-08

    Call ReadU24        '��ȡU24����
    Call ReadVDRB       '��VDR����
    Call ReadVDRM
    Call ReadAPL        '�������б�
    
    WatchName(0) = "Chief Engineer"
    WatchName(1) = "2nd Engineer"
    WatchName(3) = "3rd Engineer"
    WatchName(2) = "4th Engineer"
    WatchName(4) = "nobody"
    
    Watch1 = Val(GetSetting("RDMS System", "Duty", "Watch1", "0"))
    Watch2 = Val(GetSetting("RDMS System", "Duty", "Watch2", "4"))
    DealyW1 = Val(GetSetting("RDMS System", "Duty", "DealyW1", "1"))
    DealyW2 = Val(GetSetting("RDMS System", "Duty", "DealyW2", "3"))
    DealyW3 = Val(GetSetting("RDMS System", "Duty", "DealyW3", "2"))
    
    FormMain.ListSlaveSave.Clear
    
On Error Resume Next    '������
     If FormMain.MSComm1.PortOpen = False Then FormMain.MSComm1.PortOpen = True      '090706
     If FormMain.MSComm2.PortOpen = False Then FormMain.MSComm2.PortOpen = True      '090706
     
If MeMain = False Then Exit Sub
With FormMain   '��ں�Ϊ16��ʾ��ͨ����ʹ��

    If .MSCommK32.PortOpen = False And .MSCommK32.CommPort <> 16 Then .MSCommK32.PortOpen = True
    S = S & "  K32/COM" & .MSCommK32.CommPort & "/"
    If .MSCommK32.PortOpen = True Then S = S & .MSCommK32.Settings
    
    If .MSCommAT24.PortOpen = False And .MSCommAT24.CommPort <> 16 Then .MSCommAT24.PortOpen = True
    S = S & "  AT24/COM" & .MSCommAT24.CommPort & "/"
    If .MSCommAT24.PortOpen = True Then S = S & .MSCommAT24.Settings
    
    If .MSCommKQ16.PortOpen = False And .MSCommKQ16.CommPort <> 16 Then .MSCommKQ16.PortOpen = True
    S = S & "  KQ16/COM" & .MSCommKQ16.CommPort & "/"
    If .MSCommKQ16.PortOpen = True Then S = S & .MSCommKQ16.Settings
    
    If .MSCommQ32.PortOpen = False And .MSCommQ32.CommPort <> 16 Then .MSCommQ32.PortOpen = True
    S = S & "  Q32/COM" & .MSCommQ32.CommPort & "/"
    If .MSCommQ32.PortOpen = True Then S = S & .MSCommQ32.Settings
    
    If .MSCommEMR.PortOpen = False And .MSCommEMR.CommPort <> 16 Then .MSCommEMR.PortOpen = True
    S = S & "  EMR/COM" & .MSCommEMR.CommPort & "/"
    If .MSCommEMR.PortOpen = True Then S = S & .MSCommEMR.Settings
    
    If .MSCommU24.PortOpen = False And .MSCommU24.CommPort Then .MSCommU24.PortOpen = True
    S = S & "  U24/COM" & .MSCommU24.CommPort & "/"
    If .MSCommU24.PortOpen = True Then S = S & .MSCommU24.Settings
    
    If .MSCommVDR.PortOpen = False And .MSCommVDR.CommPort <> 16 Then .MSCommVDR.PortOpen = True
    S = S & "  VDR/COM" & .MSCommVDR.CommPort & "/"
    If .MSCommVDR.PortOpen = True Then S = S & .MSCommVDR.Settings
    
    .LabComSet.Caption = S
End With

End Sub

Public Sub SetGroup()   '��������ʼ��
Dim i As Integer
    '***************************************************************
    '��ʾҳ������
    PageGroup(1).Name = "M/E (P)"
    PageGroup(2).Name = "M/E (S)"
    PageGroup(3).Name = "G/E"
    PageGroup(4).Name = "S/G & M.A.C"
    PageGroup(5).Name = "BOILER & INC"
    PageGroup(6).Name = "OTHER"
    PageGroup(7).Name = "LEVEL"
    PageGroup(8).Name = ""
    PageGroup(1).Disp = True
    PageGroup(2).Disp = True
    PageGroup(3).Disp = True
    PageGroup(4).Disp = True
    PageGroup(5).Disp = True
    PageGroup(6).Disp = True
    PageGroup(7).Disp = True
    PageGroup(8).Disp = False
    '***************************************************************
    
    With FormMain           'ˢ�½���
    For i = 1 To 8
        .PicGroup(i).BorderStyle = 0
        .PicGroup(i).Move 0, 0, 12975, 10455
        If i <> PageNow Then .PicGroup(i).Visible = False
        .ComGroup(i).Caption = PageGroup(i).Name & " (F" & i & ")"
        .MnuGroup(i).Caption = PageGroup(i).Name
        If PageGroup(i).Disp = False Then
            .ComGroup(i).Visible = False
            .MnuGroup(i).Visible = False
        End If
    Next i
    .PicK32.Move 0, 0, 12975, 10455
    .PicAT24.Move 0, 0, 12975, 10455
    .PicComm.Move 0, 0, 12975, 10455
    
    End With
End Sub

'��ȡϵͳ����
Private Sub ReadSys()
Dim i As Integer, j As Integer
    Call File2List(App.Path & "\SetSys.ini", FormMain.ListTemp)
    
    For i = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = i
        Call Str2Array(FormMain.ListTemp.Text, ":")
        If UBound(OutStr) = 1 Then
        Select Case OutStr(0)
            Case "SysName         "
                SysName = OutStr(1)
            '����ʧ��������
            Case "MaxBinCallFail  "
                For j = 0 To 9
                    BinCall(j).MaxCall = FanWei(Val(OutStr(1)), 0, 200)
                Next j
            Case "MaxMoniCallFail "
                For j = 0 To 9
                    MoniCall(j).MaxCall = FanWei(Val(OutStr(1)), 0, 200)
                Next j
            Case "MaxKQCallFail   "
                KQCall.MaxCall = FanWei(Val(OutStr(1)), 0, 200)
            Case "MaxQ32CallFail  "
                Q32Call.MaxCall = FanWei(Val(OutStr(1)), 0, 200)
            Case "MaxEMRCallFail  "
                For j = 0 To 15
                    EMRCall(j).MaxCall = FanWei(Val(OutStr(1)), 0, 200)
                Next j
            Case "MaxU24CallFail  "
                U24Call.MaxCall = FanWei(Val(OutStr(1)), 0, 200)
            '����̨�ŷ�Χ
            Case "BinCallSta      "
                BinCallSta = FanWei(Val(OutStr(1)), 0, 9)
            Case "BinCallEnd      "
                BinCallEnd = FanWei(Val(OutStr(1)), 0, 9)
                If BinCallEnd < BinCallSta Then BinCallEnd = BinCallSta
            Case "MoniCallSta     "
                MoniCallSta = FanWei(Val(OutStr(1)), 0, 9)
            Case "MoniCallEnd     "
                MoniCallEnd = FanWei(Val(OutStr(1)), 0, 9)
                If MoniCallEnd < MoniCallSta Then MoniCallEnd = MoniCallSta
            Case "EMRCallSta      "
                EMRCallSta = FanWei(Val(OutStr(1)), 0, 9)
            Case "EMRCallEnd      "
                EMRCallEnd = FanWei(Val(OutStr(1)), 0, 9)
                If EMRCallEnd < EMRCallSta Then EMRCallEnd = EMRCallSta
            '���ںŷ���
            Case "CommK32         "
                FormMain.MSCommK32.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            Case "CommAT24        "
                FormMain.MSCommAT24.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            Case "CommKQ16        "
                FormMain.MSCommKQ16.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            Case "CommQ32         "
                FormMain.MSCommQ32.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            Case "CommEMR         "
                FormMain.MSCommEMR.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            Case "CommU24         "
                FormMain.MSCommU24.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            Case "CommVDR         "
                FormMain.MSCommVDR.CommPort = FanWei(Val(OutStr(1)), 1, 16)
            '�����ַ
            Case "IntDirM         "
                IntDirMain = OutStr(1)
            Case "IntDirS         "
                IntDirSlave = OutStr(1)
        End Select
        End If
    Next i
    MS1Call.MaxCall = 10
    MS2Call.MaxCall = 10
End Sub

'��ȡ����������
Private Sub ReadBin()
Dim i As Integer, j As Integer, k As Integer
Dim x As Integer, y As Integer, z As Integer
    For i = 0 To 9      '��ʼ��
        For j = 0 To 31
            BinData(i, j).Delay = 9
            BinData(i, j).Group = 0
            BinData(i, j).Nor = False
            BinData(i, j).Name = i & "-" & Format(j, "00")
            For k = 0 To Maxdelay
                BinData(i, j).DelayA(k) = BinData(i, j).Alm
            Next k
        Next j
    Next i
    
    Call File2List(App.Path & "\SetB.ini", FormMain.ListTemp)
    For k = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = k
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) = 7 Then
            i = FanWei(Val(OutStr(0)), 0, 9)
            j = FanWei(Val(OutStr(1)), 0, 31)
            BinData(i, j).Delay = FanWei(Val(OutStr(2)), 0, UBound(BinData(i, j).DelayA))
            BinData(i, j).Group = FanWei(Val(OutStr(3)), 0, 9)
            If Val(OutStr(4)) = 0 Then
                BinData(i, j).Nor = False
            Else
                BinData(i, j).Nor = True
            End If
            If Val(OutStr(5)) = 0 Then
                BinData(i, j).Cutout = False
            Else
                BinData(i, j).Cutout = True
            End If
            BinData(i, j).CutImg = OutStr(6)
            If Len(BinData(i, j).CutImg) = 4 Then
                x = Val(Left(BinData(i, j).CutImg, 1))
                y = Val(Right(BinData(i, j).CutImg, 2))
                '�Կ�������״̬ӳ�� ��Χ0-00 -- 9-31
                If x > 9 Or x < 0 Or y > 31 Or y < 0 Then BinData(i, j).CutImg = "9-99"
            Else
                BinData(i, j).CutImg = "9-99"
            End If
            BinData(i, j).Name = OutStr(7)
        End If
    Next k
End Sub

'��ȡģ��������
Private Sub ReadMoni()
Dim i As Integer, j As Integer, k As Integer, z As Integer
Dim A As Integer, B As Integer
    For i = 0 To 9      '��ʼ��
        For j = 0 To 23
            MoniData(i, j).Delay = 9
            MoniData(i, j).Group = 0
            MoniData(i, j).ADU = 4096
            MoniData(i, j).ADC = 2048
            MoniData(i, j).ADL = 0
            MoniData(i, j).DispU = 100
            MoniData(i, j).DispL = 0
            MoniData(i, j).Unit = " %"
            MoniData(i, j).FmtS = "00.0"
            MoniData(i, j).AlmU = 80
            MoniData(i, j).AlmL = 20
            MoniData(i, j).SFU = 4096
            MoniData(i, j).SFL = 0
            MoniData(i, j).Cutout = False
            MoniData(i, j).CutImg = "9-99"
            MoniData(i, j).Name = "M" & i & "-" & Format(j, "00")
            For k = 0 To Maxdelay
                MoniData(i, j).DelayV(k) = MoniData(i, j).Value
            Next k
        Next j
    Next i
    
    Call File2List(App.Path & "\SetM.ini", FormMain.ListTemp)    '��ȡ�����ļ�
    For k = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = k
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) = 16 Then
            i = FanWei(Val(OutStr(0)), 0, 9)
            j = FanWei(Val(OutStr(1)), 0, 23)
            MoniData(i, j).Delay = FanWei(Val(OutStr(2)), 0, UBound(MoniData(i, j).DelayV))
            MoniData(i, j).Group = FanWei(Val(OutStr(3)), 0, 9)
            MoniData(i, j).ADU = Val(OutStr(4))
            MoniData(i, j).ADL = Val(OutStr(5))
            MoniData(i, j).ADC = (MoniData(i, j).ADU + MoniData(i, j).ADL) \ 2
            MoniData(i, j).DispU = Val(OutStr(6))
            MoniData(i, j).DispL = Val(OutStr(7))
            For z = 0 To MoniData(i, j).Delay           '���Ի���������Ϊ��λ��ʾֵ
                If MoniData(i, j).DelayV(z) = 0 Then MoniData(i, j).DelayV(z) = MoniData(i, j).DispL
            Next z
            MoniData(i, j).Unit = OutStr(8)
            MoniData(i, j).FmtS = OutStr(9)
            MoniData(i, j).AlmU = Val(OutStr(10))
            MoniData(i, j).AlmL = Val(OutStr(11))
            MoniData(i, j).SFU = FanWei(Val(OutStr(12)), MoniData(i, j).ADU * 1.1, 9999)
            MoniData(i, j).SFL = FanWei(Val(OutStr(13)), 0, MoniData(i, j).ADL * 0.9)
            If Val(OutStr(14)) = 0 Then
                MoniData(i, j).Cutout = False
            Else
                MoniData(i, j).Cutout = True
            End If
            MoniData(i, j).CutImg = OutStr(15)
            If Len(MoniData(i, j).CutImg) = 4 Then  '�Կ�������״̬ӳ�� ��Χ0-00 -- 9-31
                A = Val(Left(MoniData(i, j).CutImg, 1))
                B = Val(Right(MoniData(i, j).CutImg, 2))
                If A > 9 Or A < 0 Or B > 31 Or B < 0 Then MoniData(i, j).CutImg = "9-99"
            Else
                MoniData(i, j).CutImg = "9-99"
            End If
            MoniData(i, j).Name = OutStr(16)
        End If
        If UBound(OutStr) = 17 Then
            i = FanWei(Val(OutStr(0)), 0, 9)
            j = FanWei(Val(OutStr(1)), 0, 23)
            MoniData(i, j).Delay = FanWei(Val(OutStr(2)), 0, UBound(MoniData(i, j).DelayV))
            MoniData(i, j).Group = FanWei(Val(OutStr(3)), 0, 9)
            MoniData(i, j).ADU = Val(OutStr(4))
            MoniData(i, j).ADL = Val(OutStr(6))
            MoniData(i, j).ADC = Val(OutStr(5))
            MoniData(i, j).DispU = Val(OutStr(7))
            MoniData(i, j).DispL = Val(OutStr(8))
            For z = 0 To MoniData(i, j).Delay           '���Ի���������Ϊ��λ��ʾֵ
                If MoniData(i, j).DelayV(z) = 0 Then MoniData(i, j).DelayV(z) = MoniData(i, j).DispL
            Next z
            MoniData(i, j).Unit = OutStr(9)
            MoniData(i, j).FmtS = OutStr(10)
            MoniData(i, j).AlmU = Val(OutStr(11))
            MoniData(i, j).AlmL = Val(OutStr(12))
            MoniData(i, j).SFU = FanWei(Val(OutStr(13)), MoniData(i, j).ADU * 1.1, 9999)
            MoniData(i, j).SFL = FanWei(Val(OutStr(14)), 0, MoniData(i, j).ADL * 0.9)
            If Val(OutStr(15)) = 0 Then
                MoniData(i, j).Cutout = False
            Else
                MoniData(i, j).Cutout = True
            End If
            MoniData(i, j).CutImg = OutStr(16)
            If Len(MoniData(i, j).CutImg) = 4 Then  '�Կ�������״̬ӳ�� ��Χ0-00 -- 9-31
                A = Val(Left(MoniData(i, j).CutImg, 1))
                B = Val(Right(MoniData(i, j).CutImg, 2))
                If A > 9 Or A < 0 Or B > 31 Or B < 0 Then MoniData(i, j).CutImg = "9-99"
            Else
                MoniData(i, j).CutImg = "9-99"
            End If
            MoniData(i, j).Name = OutStr(17)
        End If
    Next k
End Sub

'��ȡU24����
Private Sub ReadU24()
Dim i As Integer, j As Integer
Dim x As Integer, y As Integer
Dim A As Integer, B As Integer
    For i = 0 To 0      '��ʼ��
        For j = 0 To 23
            U24Data(i, j).DAU = 256
            U24Data(i, j).DAL = 0
            U24Data(i, j).DispImg = "9-99"
        Next j
    Next i

    Call File2List(App.Path & "\SetU.ini", FormMain.ListTemp)    '��ȡ�����ļ�
    For i = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = i
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) = 4 Then
            x = Val(OutStr(0))
            y = Val(OutStr(1))
            U24Data(x, y).DAU = FanWei(Val(OutStr(2)), 0, 256)
            U24Data(x, y).DAL = FanWei(Val(OutStr(3)), 0, 256)
            U24Data(x, y).DispImg = OutStr(4)
            If Len(U24Data(x, y).DispImg) = 4 Then  'ӳ�� ��Χ0-00 -- 9-23
                A = Val(Left(U24Data(x, y).DispImg, 1))
                B = Val(Right(U24Data(x, y).DispImg, 2))
                If A > 9 Or A < 0 Or B > 23 Or B < 0 Then U24Data(x, y).DispImg = "9-99"
            Else
                U24Data(x, y).DispImg = "9-99"
            End If
        End If
    Next i
End Sub

'��ȡVDRBin����
Private Sub ReadVDRB()
Dim i As Integer
    '��ʼ����ַΪ-1
    For i = 0 To UBound(VDRAddB)
        VDRAddB(i) = -1
    Next i
    Call File2List(App.Path & "\VDRB.ini", FormMain.ListTemp)    '��ȡ�����ļ�
    For i = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = i
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) >= 0 Then
            VDRAddB(i) = Val(OutStr(0))
        End If
    Next i
End Sub

'��ȡVDRMoni����
Private Sub ReadVDRM()
Dim i As Integer
    '��ʼ����ַΪ-1
    For i = 0 To UBound(VDRAddM)
        VDRAddM(i) = -1
    Next i
    Call File2List(App.Path & "\VDRM.ini", FormMain.ListTemp)    '��ȡ�����ļ�
    For i = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = i
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) >= 0 Then
            VDRAddM(i) = Val(OutStr(0))
        End If
    Next i
End Sub

'��ȡU24����
Private Sub ReadAPL()
Dim i As Integer, j As Integer
Dim x As Integer, y As Integer
Dim A As Integer, B As Integer
    For i = 1 To 7      '��ʼ��
        PageAlm(i) = False
        PageExAlm(i) = False
        For j = 0 To ListNum - 1
            PxAL(i, j).BorM = ""
            PxAL(i, j).AddX = 0
            PxAL(i, j).AddY = 0
        Next j
    Next i
    
    Call File2List(App.Path & "\SetAPL.ini", FormMain.ListTemp)    '��ȡ�����ļ�
    For i = 0 To FormMain.ListTemp.ListCount - 1
        FormMain.ListTemp.ListIndex = i
        Call Str2Array(FormMain.ListTemp.Text)
        If UBound(OutStr) = 4 Then
            x = Val(OutStr(0))
            y = Val(OutStr(1))
            PxAL(x, y).BorM = OutStr(2)
            PxAL(x, y).AddX = FanWei(Val(OutStr(3)), 0, 9)
            If PxAL(x, y).BorM = "B" Then PxAL(x, y).AddY = FanWei(Val(OutStr(4)), 0, 31)
            If PxAL(x, y).BorM = "M" Then PxAL(x, y).AddY = FanWei(Val(OutStr(4)), 0, 23)
        End If
    Next i
End Sub

'��������Χ�����޵������Զ�תΪ�����Сֵ
Public Function FanWei(Value As Integer, Min As Integer, Max As Integer) As Integer
    FanWei = Value
    If Value < Min Then FanWei = Min
    If Value > Max Then FanWei = Max
End Function

'Tb:������ݱ�
Public Sub ADODel(ADO As Adodc, Tb As String)
On Error Resume Next
Dim RsMy As New ADODB.Recordset
    RsMy.Open "Delete * From " & Tb, ADO.Recordset.ActiveConnection
End Sub

Public Sub ListAlmReport()
Dim i As Integer, j As Integer, k As Integer
Dim PN As Integer, AX As Integer, AY As Integer, Add As Integer
Dim S As String, SL As String
Dim t As String
Dim Bb As Boolean   '�ظ��ı�����
    Load FormReport
    t = Format(Now, "hh:mm:ss   ")
    FormReport.ListAlm.Clear
    FormReport.ListAdd.Clear
    For i = 1 To 7
        For j = 0 To ListNum - 1
            PN = (i - 1) * ListNum + j
            AX = PxAL(i, j).AddX
            AY = PxAL(i, j).AddY
            Select Case PxAL(i, j).BorM
            Case "B"
                Add = AX * 32 + AY
                SL = FormMain.LabCHName(Add).ToolTipText
                If SL = "ALM" Or SL = "SF" Then
                    S = t & FormMain.LabPList(PN).Caption
                    Bb = False
                    For k = 0 To FormReport.ListAlm.ListCount - 1
                        If FormReport.ListAlm.List(k) = S Then Bb = True
                    Next k
                    If Bb = False Then
                        FormReport.ListAlm.AddItem S
                        FormReport.ListAdd.AddItem "B" & Format(AX, "00") & "-" & Format(AY, "00")
                    End If
                End If
            Case "M"
                Add = AX * 24 + AY
                SL = FormMain.LabCHNameM(Add).ToolTipText
                If SL = "ALM" Or SL = "SF" Then
                    S = t & FormMain.LabPList(PN).Caption
                    Bb = False
                    For k = 0 To FormReport.ListAlm.ListCount - 1
                        If FormReport.ListAlm.List(k) = S Then Bb = True
                    Next k
                    If Bb = False Then
                        FormReport.ListAlm.AddItem S
                        FormReport.ListAdd.AddItem "M" & Format(AX, "00") & "-" & Format(AY, "00")
                    End If
                End If
            End Select
        Next j
    Next i
    RP = False
    FormReport.Show
End Sub

