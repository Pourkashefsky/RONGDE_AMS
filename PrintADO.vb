Attribute VB_Name = "PrintADO"

'*************** ��ӡ Adodc ***************
'Pp:ҳ�� (��������,�粻�����ϴδ�ӡ��ҳ��
'   ��Ҫ���¿�ʼ���� Pp = 0

Option Explicit
'��ӡ��ӡ�к�����
'ʹ��ʱ ReDim PrintCol(0 to ?) As Integer
Public PrintCol() As Integer
Public Pp As Integer       'ҳ��

'************* ��ӡ Adodc ************
'Ss:��ͷ
'PrintCol():��ӡ�к�����
'BeginRow,EndRow: ��ʼ�ͽ�����
'RNext:��ӡ���
'Page1: ����ÿҳ����
'ColWidth: ��׼�п�
Public Sub ADOPrint(ADO As Adodc, PrintCol() As Integer, Ss As String, BeginRow As Long, EndRow As Long, RNext As Integer, Page1 As Integer, ColWidth As Integer)
Dim StrX1 As Integer, StrY1 As Integer  'ԭ��λ��
Dim StrX As Integer, StrY As Integer    '���λ
Dim P As Integer        '��ҳ�ڼ���
Dim Wide As Integer     '����ܿ��
Dim Linw As Integer     '�и�
Dim Foot As String      'ҳ��
Dim FontS As Single     '�����С
Dim TLeft As Integer    '���������
Dim i As Integer, j As Integer, N As Integer, o As Integer
Static A(19) As Integer '��ӡ���п�����
        '?????? ԭ��λ�� ??????
    StrX1 = 1000: StrY1 = 1200
        '?????? �ɸ��и�,���� ??????
    Linw = 240
    Printer.FontName = "����"
    FontS = 8   '�����С
        '?????? ��������� ??????
    TLeft = 1500
    
    For i = 0 To UBound(PrintCol)
        A(i) = ColWidth         '�����п�
        '?????? �ɲ��������п� ??????
        '����: A(15)=2000
        A(0) = 2000
        Wide = Wide + A(i)      '�������ܿ��
    Next i
    
        '��ӡ���� & �»��� & �б���
    Call Print1(TLeft, 700, 12, Ss)
    StrX = StrX1: StrY = StrY1
    Printer.Line (StrX - 50, StrY - 30)-(StrX + Wide - 10, StrY - 30)
    For i = 0 To UBound(PrintCol)   '�б�ͷ
        Call Print1(StrX, StrY, FontS, ADO.Recordset(PrintCol(i)).Name)
        StrX = StrX + A(i)
    Next i
    StrY = StrY + Linw
    If BeginRow > ADO.Recordset.RecordCount - 1 Then
        MsgBox "Start error!"
        Exit Sub
    End If
    
    ADO.Recordset.MoveFirst '��ʼ��
    If BeginRow <> 0 Then
        For i = 0 To BeginRow
        ADO.Recordset.MoveNext
        Next i
    End If
    
    For j = BeginRow To EndRow Step RNext
        StrX = StrX1
        Printer.Line (StrX - 50, StrY - 30)-(StrX + Wide - 10, StrY - 30)
        P = P + 1
        
        For i = 0 To UBound(PrintCol)
            Call Print1(StrX, StrY, FontS, ADO.Recordset(PrintCol(i)).Value)
            StrX = StrX + A(i)
        Next i
        
        If P > Page1 Then       '��ҳ
            StrX = StrX1
            Printer.Line (StrX - 50, StrY + Linw)-(StrX + Wide - 10, StrY + Linw)
            StrY = StrY1
            For i = 0 To UBound(PrintCol)
                Printer.Line (StrX - 30, StrY - 30)-(StrX - 30, StrY + (Page1 + 2) * Linw)
                StrX = StrX + A(i)
            Next i
            Printer.Line (StrX - 30, StrY - 30)-(StrX - 30, StrY + (Page1 + 2) * Linw)
            Pp = Pp + 1     '��ӡҳ����
            Foot = "�� " + CStr(Pp) + "ҳ"
            Call Print1(StrX - 30 - 1000, StrY + (Page1 + 2) * Linw + 100, 10, Foot)
            
            Printer.NewPage
            P = 0
            Call Print1(TLeft, 700, 12, Ss) '��ӡ����
            StrX = StrX1: StrY = StrY1
            Printer.Line (StrX - 50, StrY - 30)-(StrX + Wide - 10, StrY - 30)
            For i = 0 To UBound(PrintCol)   '�б�ͷ
                Call Print1(StrX, StrY, FontS, ADO.Recordset(PrintCol(i)).Name)
                StrX = StrX + A(i)
            Next i
            StrX = StrX1: StrY = StrY + Linw
        Else
            i = 0
            Do Until ADO.Recordset.EOF Or i >= RNext
                ADO.Recordset.MoveNext
                i = i + 1
            Loop
            StrY = StrY + Linw
        End If
    Next j
    
    If P < Page1 Then   '�����ҳʣ�໮����
        For o = P To Page1 + 1
            StrX = StrX1
            Printer.Line (StrX - 50, StrY - 30)-(StrX + Wide - 10, StrY - 30)
            StrY = StrY + Linw
        Next
    End If
    
    StrX = StrX1: StrY = StrY1
    For i = 0 To UBound(PrintCol)
        Printer.Line (StrX - 30, StrY - 30)-(StrX - 30, StrY + (Page1 + 2) * Linw)
        StrX = StrX + A(i)
    Next i
    Printer.Line (StrX - 30, StrY - 30)-(StrX - 30, StrY + (Page1 + 2) * Linw)
    
    Pp = Pp + 1      '��ӡҳ����
    Foot = "�� " + CStr(Pp) + "ҳ"
    Call Print1(StrX - 30 - 1000, StrY + (Page1 + 2) * Linw + 100, 10, Foot)

    Printer.EndDoc  '��ӡ����
End Sub

'************** ��ӡ�ı� ***************
Public Sub Print1(x As Integer, y As Integer, FontS As Single, TXT As String)
    If Left(TXT, 1) = "." Then TXT = "0" & TXT
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.FontBold = False
    Printer.FontSize = FontS
    Printer.Print TXT
End Sub

