'//----------------------------------------------------------------------------
'// Uninstall SIM Information Lookup and Storage in WMI Class for SCCM
'// Compiled By: Barry Harriman
'// DATE: 23/06/2015
'// Version No: 1.0
'//----------------------------------------------------------------------------
'// CHANGE LOG
'// 
'// *********************************************************
'// 1.0 - BH - 23/06/2015 - Initial Script 
'// 
on error resume next

sVersionNo = "1.0"
EnableLogging = True
sCompanyName = "CSA"

Sub LogEntry(strLOG)
	If EnableLogging Then
		oLogFile.WriteLine date & " - " & time & " - " & strLOG
	End If
End Sub

'//----------------------------------------------------------------------------
'// Set Logging Information
'//----------------------------------------------------------------------------

Set sho = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")  

Set oShell = CreateObject("wscript.Shell")
Set fso = CreateObject("scripting.filesystemobject")
strTemp = sho.ExpandEnvironmentStrings("%temp%")

If EnableLogging Then
    Set oLogFile = fso.OpenTextFile(strTemp & "\SIMInfo.log", 8, True)
    LogEntry("*********************************************************")
    LogEntry("*********** Uninstall SIM LOOKUP Version : " & sVersionNo )
    LogEntry("*********************************************************")
End If

LogEntry(" Checking for Old Files.")
LogEntry("Clearing WMI Class")
'-------Dump value in WMI 
Dim wbemCimtypeString
Dim wbemCimtypeUint32 
wbemCimtypeString = 8 
wbemCimtypeUint32 = 19 
' Remove classes (of last run, if any)
Set oLocation = CreateObject("WbemScripting.SWbemLocator") 
Set oServices = oLocation.ConnectServer(,"root\cimv2") 
Set oNewObject = oServices.Get("CM_MobileBroadbandInfo") 
oNewObject.Delete_ 

' Create data class structure 
Set oDataObject = oServices.Get 
oDataObject.Path_.Class = "CM_MobileBroadbandInfo" 
oDataObject.Properties_.add "SIMICCID" , wbemCimtypeString 
oDataObject.Properties_.add "DeviceType" , wbemCimtypeString
oDataObject.Properties_.add "DeviceID" , wbemCimtypeString
oDataObject.Properties_.add "DeviceModel" , wbemCimtypeString
oDataObject.Properties_.add "DeviceManufacturer" , wbemCimtypeString
oDataObject.Properties_.add "IntefaceState" , wbemCimtypeString
oDataObject.Properties_.add "ProviderName" , wbemCimtypeString
oDataObject.Properties_.add "SubscriberID" , wbemCimtypeString
oDataObject.Properties_.add "SIMID" , wbemCimtypeString
oDataObject.Properties_.add "NumberTelephoneNo" , wbemCimtypeString
oDataObject.Properties_.add "TelephoneNumber" , wbemCimtypeString
oDataObject.Properties_.add "DateScriptRan" , wbemCimtypeString
oDataObject.Properties_("SIMICCID").Qualifiers_.add "key" , True
oDataObject.Put_ 

LogEntry("Finished Cleaning WMI Class")

const HKLM = &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

sKeyPath = "SOFTWARE\" & sCompanyName & "\MobileBroadband"
    
WriteLog "Deleting registry key path is HKLM\" & sKeyPath
    
oReg.DeleteKey HKLM,sKeyPath

LogEntry("Ending Script")
wscript.quit


