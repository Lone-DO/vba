Attribute VB_Name = "Accrual_js"
'
' Developed By Dayshawn Offutt (Lone-DO)
' Creation Date: 2/4/2020
' KTP Admin Asst. for David Wigginton
'

Sub Accrual_unhide()
Attribute Accrual_unhide.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Accrual_unhide.VB_ProcData.VB_Invoke_Func = " \n14"
    Rows("11:62").Select
    Selection.EntireRow.Hidden = False
End Sub

Sub Accrual_hide()
Attribute Accrual_hide.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020\n"
Attribute Accrual_hide.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim rng As Range, cell As Range
    Set rng = Range("C11:C52") ' Var Range
    For Each c In rng ' For Each Cell in Range Loop
        If c = 0 Then ' If Cell is empty || C = 0
            c.EntireRow.Select
            Selection.EntireRow.Hidden = True
        End If
    Next c

End Sub

