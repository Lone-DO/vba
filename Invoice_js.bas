Attribute VB_Name = "Invoice_js"
'
' Developed By Dayshawn Offutt (Lone-DO)
' Creation Date: 2/4/2020
' KTP Admin Asst. for David Wigginton
'

Sub Invoice_Input()
Attribute Invoice_Input.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Invoice_Input.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Invoice Entry").Activate
    Range("C10").Value = InputBox("Input the Invoice#")
    Range("C12").Value = InputBox("Input the Invoice Date")
    Range("B12").Value = InputBox("Input the PO#")
    Range("C14").Value = InputBox("Input the GRN")
    Range("B14").Value = InputBox("Input the Comments")
    Range("C16").Select
End Sub

Sub Invoice_Post()
Attribute Invoice_Post.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Invoice_Post.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Invoice Entry").Range("C16").Select 'Activate Proper Tab
    ExportEntries
    Sheets("Invoice Entry").Activate
    Range("C16").Select
End Sub

Function ExportEntries()
Set PendingArr = Sheets("Invoice Entry").Range("C16:C19")
Let EmptyData = 0
For Each c In PendingArr
    If c <> 0 Then
        Sheets("Invoice Entry").Activate
        Range(Cells(c.Row, 2), Cells(c.Row, 3)).Copy
        Sheets("Invoices").Activate
        Range("H2").Select
        Paste
        MoveData
    Else
    EmptyData = EmptyData + 1
    End If
Next
If EmptyData = PendingArr.Count Then
MsgBox "Please Finish Inputing Data/ Amounts accordingly"
Else
ClearEntry
End If


End Function

Private Sub MoveData()
    Dim rng As Range
    Set rng = Sheets("Invoices").Range("A12") 'Edit later to Offset from header,
    
    rng.EntireRow.Insert xlShiftDown, xlFormatFromRightOrBelow
    Range("A2:J2").Copy
    rng.Offset(rowoffset:=-1).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Function ClearEntry()
Let EntryArr = Array("C10", "C12", "B12", "C14", "B14") '"B10,B16,B17,B18,B19"
Set AccountArr = Sheets("Invoice Entry").Range("C16:C19")

For Each c In EntryArr
    Sheets("Invoice Entry").Range(c).ClearContents
Next

AccountArr.ClearContents
End Function
