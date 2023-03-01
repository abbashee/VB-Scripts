Const ForReading = 1 
Set objDictionary = CreateObject("Scripting.Dictionary") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = objFSO.createTextFile("Report.txt", 8, True)
Set objTextFile = objFSO.OpenTextFile("Machinelist.txt") 


On error resume next

Do Until objTextFile.AtEndOfStream 
strComputer  = objTextFile.Readline 
Passtr  = objTextFile1.Readline 
Set objUser1 = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")

If err.number=0 then

Set objUser = GetObject("WinNT://" & strComputer & "/Administrator,user") 





objUser.fullname "IBMADMIN" 
objUser.description " Description goes here " 
objUser.SetInfo 


WriteStuff.WriteLine(strComputer & vbtab & " Full Name and Desciption Changed")



Else

wscript.echo err.description

WriteStuff.WriteLine(strComputer & vbtab & Err.description & "Unable to set Full Name and Desciption")

err.clear
end if 


loop

Wscript.Echo "Done"
