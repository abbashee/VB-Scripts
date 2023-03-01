'Option Explicit
Dim objWMIService, objItem, colItems, strComputer,vfcorpcom,objExcel 


Set Fso = CreateObject("Scripting.FileSystemObject")

Set objFSO= CreateObject("Scripting.FileSystemObject")
Set InputFile = fso.OpenTextFile("MachineList.Txt")

set objCSV = objFSO.createtextfile("IPDns.txt")

Do While Not (InputFile.atEndOfStream)
HostName = InputFile.ReadLine


On Error Resume Next



Set objWMIService = GetObject("winmgmts:\\" & HostName & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_NetworkAdapterConfiguration",,48) 

if err=0 then
 
  

objcsv.writeline("-------------------------------------------------------------------------------")

 objcsv.writeline(HostName )
  
For Each objItem in colItems 
    
   objcsv.writeline(objItem.caption)
    
    If not isNull(objItem.IPAddress) Then
   
       objcsv.writeline( "IPAddress: " & Join(objItem.IPAddress))
    End If
If not isNull(objItem.DNSServerSearchOrder) Then
        
      objcsv.writeline( "DNS : " & Join(objItem.DNSServerSearchOrder))
    End If

          
Next
    End If


objcsv.writeline("------------Routing Table Info-----------------------")

set colItems1 = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_IP4RouteTable",,48) 
For Each objItem1 in colItems1 
   
  objcsv.writeline(   objItem1.Description )
   objcsv.writeline(  "Destination: " & objItem1.Destination )
    objcsv.writeline( "Information: " & objItem1.Information )
    objcsv.writeline(  "InstallDate: " & objItem1.InstallDate )
   objcsv.writeline( "InterfaceIndex: " & objItem1.InterfaceIndex )
    objcsv.writeline( "Mask: " & objItem1.Mask )
Next







objcsv.writeline("-----------------------------------")


Loop


wscript.echo "Done"
WSCript.Quit