Attribute VB_Name = "EMR"
Option Explicit

Public Watch1 As Integer            '值班人 0-4
Public Watch2 As Integer            '备用值班人 0-4
Public DealyW1 As Integer           '值班延时 1
Public DealyW2 As Integer           '值班延时 2
Public DealyW3 As Integer           '值班延时 3
Public TimeEAP As Boolean           'EMR时间同步
Public WatchName(4) As String       '值班人名称

Public GroupAlm(1 To 8) As ALMGroupType  '延伸面板八个组报警状态 组号为0不输出
Type ALMGroupType
    ExtALm As Boolean       '延伸报警态(有报警置位/全正常清零)
    NewAlm As Boolean       '新报警置位/发送后清零
End Type

Public AlmList(49) As AList         '报警列表
Type AList
    W As String         '类型 开关量/B 模拟量/M
    N As Integer        '台号
    C As Integer        '通道号
    
    No As Integer       '编号 = 启始地址 + 台号 * 24or32 + 通道号
    Value As String     '值
    Typ As String       '状态 发生报警/F 被确认/A 恢复正常/N
End Type
Public PiAL As Integer              '存取指针
Public Const BAdd As Integer = 1    '开关量启始地址 001-320
Public Const MAdd As Integer = 321  '模拟量启始地址 321-560

'报警列表堆栈初始化
Public Sub AlmListClr()
Dim i As Integer
    For i = 0 To UBound(AlmList)
        AlmList(i).No = 0
        AlmList(i).Value = "-----"
        AlmList(i).Typ = "A"
    Next i
    PiAL = 0
End Sub

'在报警列表堆栈中增加一条
Public Sub AlmListAdd(W As String, N As Integer, C As Integer, Typ As String)
    PiAL = PiAL + 1 '指针后移
    If PiAL > UBound(AlmList) Then PiAL = 0
    AlmList(PiAL).W = W
    AlmList(PiAL).N = N
    AlmList(PiAL).C = C
    AlmList(PiAL).Typ = Typ
    If W = "M" Then '模拟量
        AlmList(PiAL).No = MAdd + AlmList(PiAL).N * 24 + AlmList(PiAL).C
        If MoniData(N, C).SF = False Then
            AlmList(PiAL).Value = Left(Format(MoniData(N, C).Value, "00.00"), 5)
        Else
            AlmList(PiAL).Value = "-----"
        End If
    Else            '开关量
        AlmList(PiAL).No = BAdd + AlmList(PiAL).N * 32 + AlmList(PiAL).C
        AlmList(PiAL).Value = "-----"
    End If
End Sub

'报警列表堆栈组合字符串
Public Function AlmListStr() As String
Dim Pi As Integer
Dim S1 As String, S As String
    Pi = PiAL
    Do
        S1 = Format(AlmList(Pi).No, "000")
        S1 = S1 & "/" & AlmList(Pi).Value & AlmList(Pi).Typ
        S = S & S1 & ","
        Pi = Pi - 1
        If Pi < 0 Then Pi = UBound(AlmList)
    Loop Until Pi = PiAL    '走一圈
    AlmListStr = S
End Function

Public Sub GroupDeal()
Dim i As Integer, N As Integer
Dim NeedSave As Boolean
Dim S As String
Dim x As Integer, y As Integer, Ct As Boolean
On Error Resume Next

    For i = 1 To 8
        GroupAlm(i).ExtALm = False
        GroupAlm(i).NewAlm = False
    Next i
    
    For N = 0 To 9
    For i = 0 To 31
        If BinData(N, i).Group >= 1 And BinData(N, i).Group <= 8 Then
            S = FormMain.LabCHName(N * 32 + i).ToolTipText  '分组报警 开关量
            If S = "ALM" Or S = "ACKN-ALM" Then GroupAlm(BinData(N, i).Group).ExtALm = True
            If S = "ALM" Then GroupAlm(BinData(N, i).Group).NewAlm = True
        End If
    Next i
    Next N
    
    For N = 0 To 9
    For i = 0 To 23
        If MoniData(N, i).Group >= 1 And MoniData(N, i).Group <= 8 Then
            S = FormMain.LabCHNameM(N * 24 + i).ToolTipText   '分组报警 模拟量
            If S = "ALM" Or S = "ACKN-ALM" Then GroupAlm(MoniData(N, i).Group).ExtALm = True
            If S = "ALM" Then GroupAlm(MoniData(N, i).Group).NewAlm = True
        End If
    Next i
    Next N
        
End Sub

