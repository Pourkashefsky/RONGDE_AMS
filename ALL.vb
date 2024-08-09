Attribute VB_Name = "ALLNUM"
Option Explicit

Public EX_Normal As Integer         '����ͳ��
Public EX_Alm As Integer
Public EX_ACKAlm As Integer
Public EX_SF As Integer
Public EX_ACKSF As Integer
Public EX_Ct As Integer

Public EX_Alm_NJX As Integer        '�ǻ�е����������

'ͳ��
Public Sub Num_ALL()
On Error Resume Next
Dim i As Integer, N As Integer
    EX_Normal = 0: EX_Alm = 0: EX_ACKAlm = 0: EX_SF = 0: EX_ACKSF = 0: EX_Ct = 0
    For N = 0 To 9
        For i = 0 To 31
            If BinData(N, i).Group <> 9 Then    '9����Ϊ����ָʾ����ͳ��
                Select Case FormMain.LabCHName(N * 32 + i).ToolTipText
                    Case "NR", "": EX_Normal = EX_Normal + 1
                    Case "ALM": EX_Alm = EX_Alm + 1
                    Case "ACKN-ALM": EX_ACKAlm = EX_ACKAlm + 1
                    Case "SF": EX_SF = EX_SF + 1
                    Case "ACKN-SF": EX_ACKSF = EX_ACKSF + 1
                    Case "Cutout": EX_Ct = EX_Ct + 1
                End Select
            End If
        Next i
    Next N
    For N = 0 To 9
        For i = 0 To 23
            Select Case FormMain.LabCHNameM(N * 24 + i).ToolTipText
                Case "NR", "": EX_Normal = EX_Normal + 1
                Case "ALM": EX_Alm = EX_Alm + 1
                Case "ACKN-ALM": EX_ACKAlm = EX_ACKAlm + 1
                Case "SF": EX_SF = EX_SF + 1
                Case "ACKN-SF": EX_ACKSF = EX_ACKSF + 1
                Case "Cutout": EX_Ct = EX_Ct + 1
            End Select
        Next i
    Next N
    For N = 0 To 8
        Select Case FormMain.LabTYP(N).ToolTipText
            Case "NR": EX_Normal = EX_Normal + 1
            Case "ALM": EX_Alm = EX_Alm + 1
            Case "ACKN-ALM": EX_ACKAlm = EX_ACKAlm + 1
            Case "SF": EX_SF = EX_SF + 1
            Case "ACKN-SF": EX_ACKSF = EX_ACKSF + 1
        End Select
    Next N
    
    FormMain.SB.Panels(1).Text = "Current Alarms:" & EX_Alm & "          Total Alarms:" & EX_Alm + EX_ACKAlm & "          Cutout Alarms:" & EX_Ct
End Sub

'ͳ�Ʒǻ�е������(�����������Ѿ�ȷ�ϵı���)
Public Sub NumNJX_ALL()
On Error Resume Next
Dim i As Integer, N As Integer
Dim Address(11) As Integer
    'LFF 10��
    Address(0) = 180: Address(1) = 181
    Address(2) = 239: Address(3) = 240: Address(4) = 241: Address(5) = 242: Address(6) = 243
    Address(7) = 245: Address(8) = 246: Address(9) = 152
    '�� 2��
    Address(10) = 207: Address(11) = 208
    
    EX_Alm_NJX = 0
    For N = 0 To UBound(Address)
        Select Case FormMain.LabCHName(Address(N)).ToolTipText
            Case "ALM", "ACKN-ALM": EX_Alm_NJX = EX_Alm_NJX + 1
        End Select
    Next N
    
    '����ָʾ������
    For N = 0 To 9
        For i = 0 To 31
            If BinData(N, i).Group = 9 Then   '9����Ϊ����ָʾ
                Select Case FormMain.LabCHName(N * 32 + i).ToolTipText
                    Case "ALM", "ACKN-ALM": EX_Alm_NJX = EX_Alm_NJX + 1
                End Select
            End If
        Next i
    Next N
End Sub

'����ָʾ�ÿ��������Ƿ�Ϊ��е������
Public Function IsJX(N As Integer, i As Integer) As Boolean
Dim D As Integer
Dim Address(11) As Integer
    IsJX = True
    'LFF 10��
    Address(0) = 180: Address(1) = 181
    Address(2) = 239: Address(3) = 240: Address(4) = 241: Address(5) = 242: Address(6) = 243
    Address(7) = 245: Address(8) = 246: Address(9) = 152
    '�� 2��
    Address(10) = 207: Address(11) = 208
    
    For D = 0 To UBound(Address)
        If N * 32 + i = Address(D) Then IsJX = False
    Next D
    If BinData(N, i).Group = 9 Then IsJX = False    '9����Ϊ����ָʾ
End Function
