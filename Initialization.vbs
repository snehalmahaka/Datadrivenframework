Dim objuft

Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("E:\DataDrivenFramework\Driver\Driver4")
objuft.Test.Run
objuft.Test.Close
objuft.quit
set objuft=nothing