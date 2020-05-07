Attribute VB_Name = "Public_js"
Public Sub UnlockSheet()
Attribute UnlockSheet.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute UnlockSheet.VB_ProcData.VB_Invoke_Func = " \n14"
ActiveSheet.Unprotect
End Sub

Public Sub LockSheet()
Attribute LockSheet.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute LockSheet.VB_ProcData.VB_Invoke_Func = " \n14"
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Public Sub Copy(rng)
    Range(rng).Select
    Selection.Copy
End Sub

Public Sub Paste()
Attribute Paste.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Paste.VB_ProcData.VB_Invoke_Func = " \n14"
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End Sub


