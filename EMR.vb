Attribute VB_Name = "EMR"
Option Explicit

Public Watch1 As Integer            'ֵ���� 0-4
Public Watch2 As Integer            '����ֵ���� 0-4
Public DealyW1 As Integer           'ֵ����ʱ 1
Public DealyW2 As Integer           'ֵ����ʱ 2
Public DealyW3 As Integer           'ֵ����ʱ 3
Public TimeEAP As Boolean           'EMRʱ��ͬ��
Public WatchName(4) As String       'ֵ��������

Public GroupAlm(1 To 8) As ALMGroupType  '�������˸��鱨��״̬ ���Ϊ0�����
Type ALMGroupType
    ExtALm As Boolean       '���챨��̬(�б�����λ/ȫ��������)
    NewAlm As Boolean       '�±�����λ/���ͺ�����
End Type

Public AlmList(49) As AList         '�����б�
Type AList
    W As String         '���� ������/B ģ����/M
    N As Integer        '̨��
    C As Integer        'ͨ����
    
    No As Integer       '��� = ��ʼ��ַ + ̨�� * 24or32 + ͨ����
    Value As String     'ֵ
    Typ As String       '״̬ ��������/F ��ȷ��/A �ָ�����/N
End Type
Public PiAL As Integer              '��ȡָ��
Public Const BAdd As Integer = 1    '��������ʼ��ַ 001-320
Public Const MAdd As Integer = 321  'ģ������ʼ��ַ 321-560

'�����б��ջ��ʼ��
Public Sub AlmListClr()
Dim i As Integer
    For i = 0 To UBound(AlmList)
        AlmList(i).No = 0
        AlmList(i).Value = "-----"
        AlmList(i).Typ = "A"
    Next i
    PiAL = 0
End Sub

'�ڱ����б��ջ������һ��
Public Sub AlmListAdd(W As String, N As Integer, C As Integer, Typ As String)
    PiAL = PiAL + 1 'ָ�����
    If PiAL > UBound(AlmList) Then PiAL = 0
    AlmList(PiAL).W = W
    AlmList(PiAL).N = N
    AlmList(PiAL).C = C
    AlmList(PiAL).Typ = Typ
    If W = "M" Then 'ģ����
        AlmList(PiAL).No = MAdd + AlmList(PiAL).N * 24 + AlmList(PiAL).C
        If MoniData(N, C).SF = False Then
            AlmList(PiAL).Value = Left(Format(MoniData(N, C).Value, "00.00"), 5)
        Else
            AlmList(PiAL).Value = "-----"
        End If
    Else            '������
        AlmList(PiAL).No = BAdd + AlmList(PiAL).N * 32 + AlmList(PiAL).C
        AlmList(PiAL).Value = "-----"
    End If
End Sub

'�����б��ջ����ַ���
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
    Loop Until Pi = PiAL    '��һȦ
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
            S = FormMain.LabCHName(N * 32 + i).ToolTipText  '���鱨�� ������
            If S = "ALM" Or S = "ACKN-ALM" Then GroupAlm(BinData(N, i).Group).ExtALm = True
            If S = "ALM" Then GroupAlm(BinData(N, i).Group).NewAlm = True
        End If
    Next i
    Next N
    
    For N = 0 To 9
    For i = 0 To 23
        If MoniData(N, i).Group >= 1 And MoniData(N, i).Group <= 8 Then
            S = FormMain.LabCHNameM(N * 24 + i).ToolTipText   '���鱨�� ģ����
            If S = "ALM" Or S = "ACKN-ALM" Then GroupAlm(MoniData(N, i).Group).ExtALm = True
            If S = "ALM" Then GroupAlm(MoniData(N, i).Group).NewAlm = True
        End If
    Next i
    Next N
        
End Sub

