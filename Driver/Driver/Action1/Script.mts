'Datatable.AddSheet "Module"
'Datatable.ImportSheet "E:\KeywordDrivenFramework\Organizer\organizer.xlsx",1,"Module"
mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For i = 1 To mrowcount Step 1

Datatable.SetCurrentRow(i)

Modexe=Datatable("ModuleExe","Action1")

'msgbox Modexe
If Modexe="Y" Then

    Modid=Datatable("ModuleID","Action1")

msgbox Modid

    trowcount=datatable.GetSheet("Action2").GetRowCount

msgbox trowcount

For j=1 To trowcount Step 1
    Datatable.SetCurrentRow(j)
If Modid=Datatable("ModuleID","Action2") and Datatable("TestCaseExe","Action2")="Y" then
    testcaseid=Datatable("TestCaseID","Action2")
msgbox testcaseid
 tsrowcount=Datatable.Getsheet("Action3").GetRowCount
 msgbox tsrowcount
 For k = 1 to tsrowcount Step 1 
 datatable.SetCurrentRow(k)
 If testcaseid=Datatable("
 	
 Next

End If

Next




End If







Next
