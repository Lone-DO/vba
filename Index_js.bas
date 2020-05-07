Attribute VB_Name = "Index_js"
'
'Developed by Dayshawn Offutt (Lone-DO)
'Creattion Date: 2/3/20
'


Sub Index_Vendor_Delete()
Attribute Index_Vendor_Delete.VB_Description = "' Developed By Dayshawn Offutt (Lone-DO)\n' January - February 2020"
Attribute Index_Vendor_Delete.VB_ProcData.VB_Invoke_Func = " \n14"
    Set Target = Range("C18") 'This Targets the selection Cell
    Let first = "B10" 'This Targets the Header of the Vender Range
    Set last = Range(first).End(xlDown) 'Dynamically Finds the last active cell under Vendor Range
    Set Search = Range(first, last) 'Concatenate Ranges
    
    ' If Target (Selection) has't been chosen
    If Target = 0 Or Target = "" Then
        MsgBox ("Error, Please choose a Vendor to remove in the Drop Down")
    Else 'If Selected, then verify to remove Target
        Let Answer = MsgBox("Would you like to remove " & Target & "?", vbYesNo + vbQuestion)
        If Answer = vbYes Then 'If User accepts, Delete, Else Do Nothing
            Search.Find(Target).Delete (xlShiftUp)
        End If
    End If
    ' Clear Dropdown Selection
    Target.Select
    ActiveCell.ClearContents

End Sub

