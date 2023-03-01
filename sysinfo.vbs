Option Explicit
Dim objFSO, objExcel, objWMIService, colItems, objItem, localserver ,StrStatus, colPingedComputers ,objComputer ,objCSV 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1
localserver ="."
Dim objDomain, fso, tsInputFile, strLine, arrInput, strComputer, introw, FileLoc
introw=2
Set fso = CreateObject("Scripting.FileSystemObject")
set objCSV = FSO.createtextfile("Report.txt")
Set tsInputFile = fso.OpenTextFile("MachineList.Txt", ForReading, False)

on error resume next

Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add
objExcel.Visible = True

objExcel.Cells(1, 1).Value = "Machine Name"
objExcel.Cells(1, 2).Value = "OS"
objExcel.Cells(1, 3).Value = "OS Version"
objExcel.Cells(1, 4).Value = "Registered User"  
objExcel.Cells(1, 5).Value = "Serial Number"  
objExcel.Cells(1, 6).Value = "CSD Version"  
objExcel.Cells(1, 7).Value = "Description"  
objExcel.Cells(1, 8).Value = "Last Boot Up Time"  
objExcel.Cells(1, 9).Value = "Local Date Time"  
objExcel.Cells(1, 10).Value = "Organization"  
objExcel.Cells(1, 11).Value = "Domain"  
 
objExcel.Cells(1, 12).Value = "Model"  
objExcel.Cells(1, 13).Value = "Number Of Processors"  
objExcel.Cells(1, 14).Value = "Primary Owner Name"
objExcel.Cells(1, 15).Value = "System Type"
objExcel.Cells(1, 16).Value = "Total Physical Memory"

objExcel.Cells(1, 17).Value = "User Name"
objExcel.Cells(1, 18).Value = "Caption"

objExcel.Cells(1, 17).Value = "Manufacturer"

'objExcel.Cells(1, 18).Value = "Name"

objExcel.Cells(1, 18).Value = "Release Date"
objExcel.Cells(1, 19).Value = "Serial Number"
objExcel.Cells(1, 20).Value = "SMBIOS BIOS Version"
objExcel.Cells(1, 21).Value = "BIOS Version"
objExcel.Cells(1, 22).Value = "CPU cores"
introw=2

While Not tsInputFile.AtEndOfStream
   strComputer = tsInputFile.ReadLine
  



Set objWMIService = GetObject ("winmgmts:\\" & localserver & "\root\cimv2")



Set colPingedComputers = objWMIService.ExecQuery ("Select * from Win32_PingStatus Where Address = '" & strComputer &"'")
	For Each objComputer in colPingedComputers
    		If objComputer.StatusCode = 0 Then
        		
		 
			StrStatus=0
    		Else
        	
			StrStatus=1
 
   		End If
         Next



 Scanpc (strComputer)





introw = introw + 1
Wend

tsInputFile.Close
WScript.Echo "Finished"
WScript.Quit(0)


Sub scanpc(strComputer)

if StrStatus=0 then

 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
 For Each objItem in colItems
 objExcel.Cells(introw, 1).Value = objItem.CSName
 objExcel.Cells(introw, 2).Value = objItem.Caption
 objExcel.Cells(introw, 3).Value = objItem.Version
 objExcel.Cells(introw, 4).Value = objItem.RegisteredUser
 objExcel.Cells(introw, 5).Value = objItem.SerialNumber
 objExcel.Cells(introw, 6).Value = objItem.CSDVersion
 objExcel.Cells(introw, 7).Value = objItem.Description
 objExcel.Cells(introw, 8).Value = objItem.LastBootUpTime
 objExcel.Cells(introw, 9).Value = objItem.LocalDateTime
 objExcel.Cells(introw, 10).Value = objItem.Organization
 Next
 Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
 For Each objItem in colItems
 objExcel.Cells(introw, 11).Value = objItem.Domain

 objExcel.Cells(introw, 12).Value = objItem.Model
objExcel.Cells(introw, 13).Value = objItem.NumberOfProcessors
objExcel.Cells(introw, 14).Value = objItem.PrimaryOwnerName
objExcel.Cells(introw, 15).Value = objItem.SystemType
objExcel.Cells(introw, 16).Value = (objItem.TotalPhysicalMemory /1024) & " MB"
'objExcel.Cells(introw, 17).Value = objItem.UserName
 Next
 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS")
 For Each objItem in colItems
'objExcel.Cells(introw, 18).Value = objItem.Caption
 objExcel.Cells(introw, 17).Value = objItem.Manufacturer
 'objExcel.Cells(introw, 18).Value = objItem.Name
 objExcel.Cells(introw, 18).Value = objItem.ReleaseDate
 objExcel.Cells(introw, 19).Value = objItem.SerialNumber
 objExcel.Cells(introw, 20).Value = objItem.SMBIOSBIOSVersion
 objExcel.Cells(introw, 21).Value = objItem.Version
 Next

 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
 For Each objItem in colItems
objExcel.Cells(introw, 22).Value = objItem.NumberOfCores
 
 Next


else
objExcel.Cells(introw, 1).Value = strComputer + "---> Did not respond"

end if
 objExcel.Range("A1:Z1").Select
 objExcel.Selection.Font.ColorIndex = 11
 objExcel.Selection.Font.Bold = True
 objExcel.Cells.EntireColumn.AutoFit
if err then

objCSV.writeline( strComputer + "---->" + err.description)

err.clear

end if 

end Sub 
