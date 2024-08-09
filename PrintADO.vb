Attribute VB_Name = "PrintADO"

'*************** 打印 Adodc ***************
'Pp:页码 (公共变量,如不接着上次打印的页数
'   而要重新开始设置 Pp = 0

Option Explicit
'打印打印列号数组
'使用时 ReDim PrintCol(0 to ?) As Integer
Public PrintCol() As Integer
Public Pp As Integer       '页码

'************* 打印 Adodc ************
'Ss:表头
'PrintCol():打印列号数组
'BeginRow,EndRow: 开始和结束行
'RNext:打印间隔
'Page1: 定义每页行数
'ColWidth: 标准列宽
Public Sub ADOPrint(ADO As Adodc, PrintCol() As Integer, Ss As String, BeginRow As Long, EndRow As Long, RNext As Integer, Page1 As Integer, ColWidth As Integer)
Dim StrX1 As Integer, StrY1 As Integer  '原点位置
Dim StrX As Integer, StrY As Integer    '单项定位
Dim P As Integer        '本页第几项
Dim Wide As Integer     '表格总宽度
Dim Linw As Integer     '行高
Dim Foot As String      '页脚
Dim FontS As Single     '字体大小
Dim TLeft As Integer    '标题横坐标
Dim i As Integer, j As Integer, N As Integer, o As Integer
Static A(19) As Integer '打印的列宽数组
        '?????? 原点位置 ??????
    StrX1 = 1000: StrY1 = 1200
        '?????? 可改行高,字体 ??????
    Linw = 240
    Printer.FontName = "宋体"
    FontS = 8   '字体大小
        '?????? 标题横坐标 ??????
    TLeft = 1500
    
    For i = 0 To UBound(PrintCol)
        A(i) = ColWidth         '定义列宽
        '?????? 可插入特殊列宽 ??????
        '例如: A(15)=2000
        A(0) = 2000
        Wide = Wide + A(i)      '计算表格总宽度
    Next i
    
        '打印标题 & 下划线 & 列标题
    Call Print1(TLeft, 700, 12, Ss)
    StrX = StrX1: StrY = StrY1
    Printer.Line (StrX - 50, StrY - 30)-(StrX + Wide - 10, StrY - 30)
    For i = 0 To UBound(PrintCol)   '列标头
        Call Print1(StrX, StrY, FontS, ADO.Recordset(PrintCol(i)).Name)
        StrX = StrX + A(i)
    Next i
    StrY = StrY + Linw
    If BeginRow > ADO.Recordset.RecordCount - 1 Then
        MsgBox "Start error!"
        Exit Sub
    End If
    
    ADO.Recordset.MoveFirst '启始行
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
        
        If P > Page1 Then       '换页
            StrX = StrX1
            Printer.Line (StrX - 50, StrY + Linw)-(StrX + Wide - 10, StrY + Linw)
            StrY = StrY1
            For i = 0 To UBound(PrintCol)
                Printer.Line (StrX - 30, StrY - 30)-(StrX - 30, StrY + (Page1 + 2) * Linw)
                StrX = StrX + A(i)
            Next i
            Printer.Line (StrX - 30, StrY - 30)-(StrX - 30, StrY + (Page1 + 2) * Linw)
            Pp = Pp + 1     '打印页角码
            Foot = "第 " + CStr(Pp) + "页"
            Call Print1(StrX - 30 - 1000, StrY + (Page1 + 2) * Linw + 100, 10, Foot)
            
            Printer.NewPage
            P = 0
            Call Print1(TLeft, 700, 12, Ss) '打印标题
            StrX = StrX1: StrY = StrY1
            Printer.Line (StrX - 50, StrY - 30)-(StrX + Wide - 10, StrY - 30)
            For i = 0 To UBound(PrintCol)   '列标头
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
    
    If P < Page1 Then   '在最后页剩余划空行
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
    
    Pp = Pp + 1      '打印页角码
    Foot = "第 " + CStr(Pp) + "页"
    Call Print1(StrX - 30 - 1000, StrY + (Page1 + 2) * Linw + 100, 10, Foot)

    Printer.EndDoc  '打印结束
End Sub

'************** 打印文本 ***************
Public Sub Print1(x As Integer, y As Integer, FontS As Single, TXT As String)
    If Left(TXT, 1) = "." Then TXT = "0" & TXT
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.FontBold = False
    Printer.FontSize = FontS
    Printer.Print TXT
End Sub

