
Dim objWMIService, objItem, colItems, strComputer,vfcorpcom,objExcel 


Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

objExcel.Workbooks.Add

intRow = 2



objExcel.Cells(1, 1).Value = "Server Name"
objExcel.Cells(1, 2).Value = "Drives"
objExcel.Cells(1, 3).Value = "Total Space (GB)"
objExcel.Cells(1, 4).Value = "Free Space (GB)"
objExcel.Cells(1, 5).Value = "% of Free Space"
 

Set Fso = CreateObject("Scripting.FileSystemObject")
Set InputFile = fso.OpenTextFile("MachineList.Txt")


Do While Not (InputFile.atEndOfStream)
HostName = InputFile.ReadLine


On Error Resume Next

strComputer=HostName



Set objWMIService = GetObject _
("winmgmts:\\" & strComputer & "\root\cimv2")

if  err.description=""  then



Set colItems = objWMIService.ExecQuery _
("Select * from Win32_LogicalDisk  WHERE DriveType=3 ")



For Each objItem in colItems
objExcel.Cells(intRow, 1).Value =  strComputer
objExcel.Cells(intRow, 2).Value =  objItem.Name
objExcel.Cells(intRow, 3).Value =  round((objItem.Size /1073741824),2)
objExcel.Cells(intRow, 4).Value = round((objItem.FreeSpace /1073741824),2)
objExcel.Cells(intRow, 5).Value = INT((objItem.FreeSpace / objItem.Size) * 1000)/10 & " %"


intRow = intRow + 1

Next


else

objExcel.Cells(intRow, 1).Value =  strComputer
objExcel.Cells(intRow, 2).Value = err.description



intRow = intRow + 1


err.clear
end if

Loop

Set objRange = objExcel.Range("A1","E1")
objRange.Font.Bold = TRUE

Set objRange = objExcel.Range("A1","E1")
objRange.Font.Size = 12

Set objRange = objExcel.Range("A1","E1")
objRange.Interior.ColorIndex = 40

'Set objRange = objExcel.ActiveCell.EntireColumn
'objRange.AutoFit()

objExcel.Range("A1:A25").Select
 'objExcel.Selection.Font.ColorIndex = 11
 'objExcel.Selection.Font.Bold = True
 objExcel.Cells.EntireColumn.AutoFit





WSCript.Echo "Done"






