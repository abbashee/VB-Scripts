
Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")

RegKey= "PointAndPrint"
strKeyPath = "SOFTWARE\Policies\Microsoft\Windows NT\Printers"

oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath,regkey, arrSubKeys

if Isnull(arrSubKeys) then 
 wsh.echo "not Found:", sValue
else

wsh.echo " Found:", sValue


End if 
