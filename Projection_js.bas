Attribute VB_Name = "Projection_js"
'
' Developed By Dayshawn Offutt (Lone-DO)
' Creation Date: 2/4/2020
' KTP Admin Asst. for David Wigginton
'

Sub PjxClearData()
Attribute PjxClearData.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute PjxClearData.VB_ProcData.VB_Invoke_Func = " \n14"
UnlockSheet

    CellGroups = Array("F11:K15", "F18:K24", "F27:K30", "F33:K36", "F39:K42", "F45:K46", "F49:K51", "F54:K55", "F58:K60", "F63:K65")
    Answer = MsgBox("WARNING: UNRECOVERABLE DATA; Are you sure you wish to clear this Sheet?", vbYesNo + vbQuestion)
    If Answer = vbYes Then
        For Each rng In CellGroups
        Range(rng).ClearContents
        Next
    End If
    ActiveWindow.ScrollRow = 1
    Range("D9").Select
    
LockSheet
End Sub
Sub Pjx_First()
Attribute Pjx_First.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Pjx_First.VB_ProcData.VB_Invoke_Func = " \n14"
    Export 6
End Sub
Sub Pjx_Second()
Attribute Pjx_Second.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Pjx_Second.VB_ProcData.VB_Invoke_Func = " \n14"
    Export 7
End Sub
Sub Pjx_Third()
Attribute Pjx_Third.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Pjx_Third.VB_ProcData.VB_Invoke_Func = " \n14"
    Export 8
End Sub
Sub Pjx_Fourth()
Attribute Pjx_Fourth.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Pjx_Fourth.VB_ProcData.VB_Invoke_Func = " \n14"
    Export 9
End Sub
Sub Pjx_Final()
Attribute Pjx_Final.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Pjx_Final.VB_ProcData.VB_Invoke_Func = " \n14"
    Export 10
End Sub
Private Sub Export(cCol)
Let Target = "B11"
Let PnL = "B11:B66"
Let cRow = 11
Set Current = Cells(cRow, cCol)
'Variables & Declarations above
UnlockSheet
Copy PnL
Current.Select
Validate cCol, Current
Range(Target).Select
LockSheet
End Sub
Function Cycle(n)
'Loop through columns, if one is full, move to next
    If (n >= 11) Then
        MsgBox ("Pjx Complete")
    ElseIf Cells(11, n) = "" Then
        'Range(Target).RemoveSubtotal
        Copy "B11:B66"
        Cells(11, n).Select
        Paste
    Else
        Cycle (n + 1)
    End If
End Function
Function Validate(cCol, Current)
    If Current <> "" Then
        If MsgBox("Overwrite? " & Cells(9, cCol), vbYesNo + vbQuestion) = vbYes Then
            Paste
        End If
    Else
        Paste
    End If
End Function
Private Sub PjxExportCycle()
UnlockSheet
    ' Programmed by Dayshawn Offutt (Lone-DO)
    Let Target = "B11"
    Let SystemPjx = "B11:B66"
    Let cCol = 6
    Let cRow = 11
    
    Cycle cCol
    Range(Target).Select
    'AddSubtotal
LockSheet
End Sub
