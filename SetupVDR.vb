Private Sub ComAdd_Click(Index As Integer)
    Select Case Index
    Case 0
        ListBin.AddItem ComboBin.Text, ListBin.ListIndex
    Case 1
        ListMoni.AddItem ComboMoni.Text, ListMoni.ListIndex
    End Select
End Sub

Private Sub ComClose_Click()
    Unload Me
End Sub

Private Sub ComDel_Click(Index As Integer)
    Select Case Index
    Case 0
        ListBin.RemoveItem ListBin.ListIndex
    Case 1
        ListMoni.RemoveItem ListMoni.ListIndex
    End Select
End Sub

Private Sub ComMDW_Click(Index As Integer)
Dim S As String
    Select Case Index
    Case 0
        If ListBin.ListIndex < ListBin.ListCount - 1 Then
            S = ListBin.List(ListBin.ListIndex)
            ListBin.List(ListBin.ListIndex) = ListBin.List(ListBin.ListIndex + 1)
            ListBin.List(ListBin.ListIndex + 1) = S
            ListBin.ListIndex = ListBin.ListIndex + 1
        End If
    Case 1
        If ListMoni.ListIndex < ListMoni.ListCount - 1 Then
            S = ListMoni.List(ListMoni.ListIndex)
            ListMoni.List(ListMoni.ListIndex) = ListMoni.List(ListMoni.ListIndex + 1)
            ListMoni.List(ListMoni.ListIndex + 1) = S
            ListMoni.ListIndex = ListMoni.ListIndex + 1
        End If
    End Select
End Sub

Private Sub ComMUP_Click(Index As Integer)
Dim S As String
    Select Case Index
    Case 0
        If ListBin.ListIndex > 0 Then
            S = ListBin.List(ListBin.ListIndex)
            ListBin.List(ListBin.ListIndex) = ListBin.List(ListBin.ListIndex - 1)
            ListBin.List(ListBin.ListIndex - 1) = S
            ListBin.ListIndex = ListBin.ListIndex - 1
        End If
    Case 1
        If ListMoni.ListIndex > 0 Then
            S = ListMoni.List(ListMoni.ListIndex)
            ListMoni.List(ListMoni.ListIndex) = ListMoni.List(ListMoni.ListIndex - 1)
            ListMoni.List(ListMoni.ListIndex - 1) = S
            ListMoni.ListIndex = ListMoni.ListIndex - 1
        End If
    End Select
End Sub

Private Sub ComSave_Click()
    Call List2File(ListBin, App.Path & "\VDRB.ini")
    Call List2File(ListMoni, App.Path & "\VDRM.ini")
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
    For i = 0 To 9
    For j = 0 To 31
        ComboBin.AddItem Format(i * 32 + j, "000") & vbTab & BinData(i, j).Name
    Next j
    Next i
    For i = 0 To 9
    For j = 0 To 23
        ComboMoni.AddItem Format(i * 24 + j, "000") & vbTab & MoniData(i, j).Name
    Next j
    Next i
    
    i = 0
    ListBin.Clear
    Do While VDRAddB(i) >= 0
        ListBin.AddItem Format(VDRAddB(i), "000") & vbTab & BinData(VDRAddB(i) \ 32, VDRAddB(i) Mod 32).Name
        i = i + 1
    Loop
    ListBin.ListIndex = 0
    
    i = 0
    ListMoni.Clear
    Do While VDRAddM(i) >= 0
        ListMoni.AddItem Format(VDRAddM(i), "000") & vbTab & MoniData(VDRAddM(i) \ 24, VDRAddM(i) Mod 24).Name
        i = i + 1
    Loop
    ListMoni.ListIndex = 0
End Sub

Private Sub ListBin_Click()
    ComboBin.ListIndex = Val(Left(ListBin.Text, 3))
End Sub

Private Sub ListMoni_Click()
    ComboMoni.ListIndex = Val(Left(ListMoni.Text, 3))
End Sub

Private Sub TextPassW_Change()
    If TextPassW.Text = Password Then
        ComSave.Enabled = True
    Else
        ComSave.Enabled = False
    End If
End Sub
