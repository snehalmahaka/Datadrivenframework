﻿'Datatable.AddSheet "Module"
'Datatable.ImportSheet "C:\Saritha\KeywordDrivenFramework\Organizer\organizer.xlsx",1,"Module"
'Datatable.ImportSheet "C:\Saritha\KeywordDrivenFramework\Organizer\organizer.xlsx",2,"Testcase"
'Datatable.ImportSheet "C:\Saritha\KeywordDrivenFramework\Organizer\organizer.xlsx",3,"TestStep"

 


Services.StartTransaction "tr1"

 

mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount
For i = 1 To mrowcount Step 1
Datatable.SetCurrentRow(i)
Modexe=Datatable("Moduleexe","Action1")
'msgbox Modexe
If Modexe="Y" Then
        Modid=Datatable("ModuleID","Action1")
        msgbox Modid
        trowcount=datatable.GetSheet("Action2").GetRowCount
        msgbox trowcount
        For j=1 To trowcount Step 1
    Datatable.SetCurrentRow(j)
    If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" then
    testcaseid=Datatable("TestcaseId","Action2")
    msgbox testcaseid
        tsrowcount=Datatable.GetSheet("Action3").GetRowCount
        msgbox tsrowcount
        For k = 1 to tsrowcount Step 1
        datatable.SetCurrentRow(k)
        If testcaseid=Datatable("TestcaseId","Action3") Then
        keyword=Datatable("Keyword","Action3")
        msgbox keyword
        Select case (keyword)
        
        Case "In"
        Call Login(john,hp)
        
        Case "ca"
        Call Closeapp()
        
        Case "oo"
        Call openorder()
        Case "uo"
        Call Updateorder()
        
        Case "lnd"
        drowcount=datatable.GetSheet("Action4").GetRowCount
        
        For l = 1 To drowcount Step 1
        	
        	datatable.SetCurrentRow(1)
        	
        	Call login(datatable("username","Action4"),datatable("password","Action4"))
        	
        	Call Closeapp()
        Next
        
        Case "ood"
        
        orrowcount=datatable.GetSheet("Action4").GetRowCount
        
        For m = 1 To orrowcount Step 1
        	datatable.SetCurrentRow(m)
            Call openorder(datatable("orderno","Action4"))
        Next
        
        End  Select
        
        End If
            
        Next
        
        
    End If
        
    Next
    
    
    
    
End If

 


    
    
    
Next


 


Services.EndTransaction "tr1"



 @@ hightlight id_;_1904626408_;_script infofile_;_ZIP::ssf14.xml_;_
 @@ hightlight id_;_2046860568_;_script infofile_;_ZIP::ssf19.xml_;_
