'Abdul Basheer
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set InputFile = objfso.OpenTextFile("MachineList.Txt")

set objTextFile = objFSO.CreateTextFile("Report.txt")

HKEY_LOCAL_MACHINE = &H80000002

keypath= "System\CurrentControlSet\Services\HTTP\Parameters\"
strKey= "EnableTrailerSupport"


'On Error Resume Next

Do While Not (InputFile.atEndOfStream)
strComputerName =  InputFile.ReadLine

set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\default:StdRegProv")

if  err.description=""  then

If objReg.EnumKey(HKEY_LOCAL_MACHINE, keypath, strkey) = 0 Then
objTextFile.writeline(  strComputername & Vbtab & "EnableTrailerSupport Key : Exists " )

objReg.GetDWORDValue HKEY_LOCAL_MACHINE,KeyPath,key,dwValue

if isnull(dwValue) then
objTextFile.writeline(  strComputername & Vbtab & "The Windows EnableTrailerSupport is:  Not defined " )
else
objTextFile.writeline(  strComputername & Vbtab & "The Windows EnableTrailerSupport is: " & dwValue)
end if





Else

objTextFile.writeline(  strComputername & Vbtab & "EnableTrailerSupport: Doesn't Exists " )



End If





Else

  objTextFile.writeline(  strComputername & Vbtab & err.description)


end if

loop

Wscript.echo "Done"