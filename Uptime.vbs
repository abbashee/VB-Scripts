On Error Resume Next

set objFSO = CreateObject("Scripting.FileSystemObject")


Set InputFile = objFSO.OpenTextFile("MachineList.Txt")

set objCSV = objFSO.createtextfile("Uptimereport.txt")

Do While Not (InputFile.atEndOfStream)

strComputer  = InputFile.ReadLine

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
 
For Each objOS in colOperatingSystems
    dtmBootup = objOS.LastBootUpTime
    dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
    dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
    objCSV.writeline( strComputer  & vbtab & dtmSystemUptime )
Next
 
Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & _
        Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
            & " " & Mid (dtmBootup, 9, 2) & ":" & _
                Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup,13, 2))
End Function