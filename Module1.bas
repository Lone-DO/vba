Attribute VB_Name = "Module1"
'
' Developed By Dayshawn Offutt (Lone-DO)
' Date: 2/7/2020
'

Sub ToggleSubtotal()
    Set AccountName = Range("H12")
    Set InvoiceDate = Range("B12")
    ' NOTE: Cannot Add subtotal to Tables... Must convert into Range/ normal cells
    
    Sheets("Invoices").Range("H12").Select
    Selection.removeSubtotal
    Selection.Sort Key1:=AccountName, Order1:=xlAscending, Key2:=InvoiceDate _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Selection.Subtotal GroupBy:=7, Function:=xlSum, TotalList:=Array(8), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
End Sub

Sub RemoveSub()
    Set DateEntered = Range("B12")
    Set InvoiceDate = Range("D12")
    Sheets("Invoices").Range("H12").Select
    
    Selection.removeSubtotal
    Selection.Sort Key1:=DateEntered, Order1:=xlAscending, Key2:=InvoiceDate _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
End Sub
