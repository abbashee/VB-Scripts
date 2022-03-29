Const ForReading = 1 
Set objDictionary = CreateObject("Scripting.Dictionary") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = objFSO.createTextFile("Report.txt")
Set objTextFile = objFSO.OpenTextFile("Machinelist.txt") 
Passtr  = "new Password here"

On error resume next

Do Until objTextFile.AtEndOfStream 
strComputer  = objTextFile.Readline 

Set objUser1 = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")

If err.number=0 then

Set objUser = GetObject("WinNT://" & strComputer & "/win5982,user")'<----Enter the Admin Account here. 





objUser.SetPassword Passtr ' <--- this will be the new admin password. 
objUser.SetInfo 


WriteStuff.WriteLine(strComputer & vbtab & " Password set")



Else


WriteStuff.WriteLine(strComputer & vbtab & Err.description & "Unable to set pass")

err.clear
end if 


loop

Wscript.Echo"Done"
