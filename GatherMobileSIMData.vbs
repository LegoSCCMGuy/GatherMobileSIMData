'//----------------------------------------------------------------------------
'// SIM Information Lookup and Storage in WMI Class for SCCM
'// Compiled By: Barry Harriman
'// DATE: 13/08/2015
'// Version No: 1.6
'//----------------------------------------------------------------------------
'// CHANGE LOG
'// 
'// *********************************************************
'// 1.1 - BH - 17/06/2015 - Initial Script 
'// 1.2 - BH - 23/06/2015 - Modified 
'// 1.3 - BH - 13/08/2015 - Modified TS
'// 1.4 - BH - 13/08/2015 - Modified HW INV
'// 1.5 - BH - 02/02/2018 - Modified to Support Windows 10 Cellular Details
'// 1.6 - BH - 29/08/2018 - Modified to add additional MDN Naming
on error resume next

sVersionNo = "1.6"
EnableLogging = True
sCompanyName = "CSA"

Sub DeleteAFile(filespec)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(filespec)
End Sub

Sub LogEntry(strLOG)
	If EnableLogging Then
		oLogFile.WriteLine date & " - " & time & " - " & strLOG
	End If
End Sub

Sub ClearOldInfo(strOldFileName)
	if fso.fileexists(strTemp & "\" & strOldFileName) then
		LogEntry("  ** File Found: " & strTemp & "\" & strOldFileName & ". Deleting File")
		DeleteAFile(strTemp & "\" & strOldFileName)
	else
		LogEntry("  " & strOldFileName & " not found.")
	end if
end sub

function LuhnDouble(intNumb)
	intNumb = intNumb * 2
	if intNumb >= 10 then intNumb = intNumb - 9 End if  
	LuhnDouble = intNumb
End function

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
    LogEntry("***********  SIM LOOKUP Version : " & sVersionNo )
    LogEntry("*********************************************************")
End If

LogEntry(" Checking for Old Files.")
ClearOldInfo("meid.txt")

DIM strCMDtoRun(5)
	strCMDtoRun(0)="cmd /C netsh mbn show interfaces >> " & strTemp & "\meid.txt"
	strCMDtoRun(1)="cmd /C netsh mbn show readyinfo interface=" & chr(34) & "Mobile Broadband Connection" & chr(34) & " >> " & strTemp & "\meid.txt"
	strCMDtoRun(2)="cmd /C netsh mbn show readyinfo interface=" & chr(34) & "Mobile Broadband Connection 2" & chr(34) & " >> " & strTemp & "\meid.txt"
	strCMDtoRun(3)="cmd /C netsh mbn show readyinfo interface=" & chr(34) & "Mobile Broadband Connection 3" & chr(34) & " >> " & strTemp & "\meid.txt"
	strCMDtoRun(4)="cmd /C netsh mbn show readyinfo interface=" & chr(34) & "Mobile Broadband" & chr(34) & " >> " & strTemp & "\meid.txt"
	strCMDtoRun(5)="cmd /C netsh mbn show readyinfo interface=" & chr(34) & "Cellular" & chr(34) & " >> " & strTemp & "\meid.txt"

LogEntry(" Running Commands.") 
for i=0 to 5
	'wscript.echo strcmdtorun(i)
	sho.run strcmdtorun(i),0,vbtrue
next

LogEntry(" Reading data from Stored Files.")

DeviceType="Not Found"
DeviceID="Not Found"
DeviceModel="Not Found"
DeviceManufacturer="Not Found"
IntefaceState="Not Found"
ProviderName="Not Found"
SubscriberID="Not Found"
SIMICCID="Not Found"
NumberTelephoneNo="Not Found"
TelephoneNumber=""
SIMID=""

LogEntry(" Checking meid.txt")
if fso.fileexists(strTemp & "\meid.txt") then
	LogEntry(" Reading data from meid.txt") 
	Set objMEIDFile = FSO.OpenTextFile(strTemp & "\meid.txt",1) 
	do until objMEIDFile.AtEndOfStream
		strMeidData=objMEIDFile.ReadLine
		'LogEntry(" ** Trace - " & strMeidData)
		intMeidDataSplit = InStr (1, strMeidData, ":", vbTextCompare)
		strSettingName = trim( left(strMeidData, intMeidDataSplit - 1))
		'LogEntry(" Data type: '" & strSettingName & "'")
		strSetting = trim( right(strMeidData, len(strMeidData) - intMeidDataSplit))
		'LogEntry(" Data: '" & strSetting & "'")

		if strSettingName="Device type" then 
			DeviceType=strSetting
		end if
		if strSettingName="Device Id" then
			DeviceID=strSetting
		end if
		if strSettingName="Model" then
			DeviceModel=strSetting
		end if
		if strSettingName="Manufacturer" then
			DeviceManufacturer=strSetting
		end if
		if strSettingName="State" then
			IntefaceState=strSetting
		end if
		if strSettingName="Provider Name" then
			ProviderName=strSetting
		end if
		if strSettingName="Subscriber Id" then
			SubscriberID=strSetting
		end if
		if strSettingName="SIM ICC Id" then
			SIMICCID=strSetting
		end if	
		if strSettingName="Number of telephone numbers" then
			NumberTelephoneNo=strSetting
		end if				
		strSettingName=""
		strSetting=""
	Loop
	LogEntry(" Finished Reading data from meid.txt")
end if
LogEntry(" Finished checking meid.txt")

LogEntry(" DeviceType='" & DeviceType & "'")
LogEntry(" DeviceID='" & DeviceID & "'")
LogEntry(" DeviceModel='" & DeviceModel & "'")
LogEntry(" DeviceManufacturer='" & DeviceManufacturer & "'")
LogEntry(" IntefaceState='" & IntefaceState & "'")
LogEntry(" ProviderName='" & ProviderName & "'")
LogEntry(" SubscriberID='" & SubscriberID & "'")
LogEntry(" SIMICCID='" & SIMICCID & "'")
LogEntry(" SIMID='" & SIMID & "'")
LogEntry(" NumberTelephoneNo='" & NumberTelephoneNo & "'")
LogEntry(" TelephoneNumber='" & TelephoneNumber & "'")
LogEntry("Calculating Check Digit")
if len(SIMICCID) > 12 Then
	CalcSimID = mid(SIMICCID,7,12)
	DIM CalcSimArray(14)
	CalcSimChk = 0
	CalcSimChkVal = 0
	for i=0 to 11
		CalcSimArray(i) = mid(CalcSimID,i+1,1)
		if i=1 then CalcSimArray(i) = LuhnDouble(CalcSimArray(i)) end if
		if i=3 then CalcSimArray(i) = LuhnDouble(CalcSimArray(i)) end if
		if i=5 then CalcSimArray(i) = LuhnDouble(CalcSimArray(i)) end if
		if i=7 then CalcSimArray(i) = LuhnDouble(CalcSimArray(i)) end if
		if i=9 then CalcSimArray(i) = LuhnDouble(CalcSimArray(i)) end if
		if i=11 then CalcSimArray(i) = LuhnDouble(CalcSimArray(i)) end if
		
		CalcSimChk = CalcSimChk * 10 + CalcSimArray(i)
		CalcSimChkVal = CalcSimChkVal + CalcSimArray(i)
	Next
	CalcSimChkVal = Right(CalcSimChkVal * 9,1)
	SIMID = CalcSimID & "-" & CalcSimChkVal
	LogEntry("SIMID: " & SIMID )
end if

LogEntry("Populating WMI Class")
'-------Dump values into WMI 
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

'------------------------------
'Add Instances to data class
Set oServices = oLocation.ConnectServer(, "root\cimv2")
Set oNewObject = oServices.Get("CM_MobileBroadbandInfo").SpawnInstance_
oNewObject.SIMICCID = SIMICCID
oNewObject.SIMID = SIMID
oNewObject.DeviceType = DeviceType
oNewObject.DeviceID = DeviceID
oNewObject.DeviceModel = DeviceModel
oNewObject.DeviceManufacturer = DeviceManufacturer
oNewObject.IntefaceState = IntefaceState
oNewObject.ProviderName = ProviderName
oNewObject.SubscriberID = SubscriberID
oNewObject.NumberTelephoneNo = NumberTelephoneNo
oNewObject.TelephoneNumber = TelephoneNumber
oNewObject.DateScriptRan = Now
oNewObject.Put_


LogEntry("Finished Populating WMI Class")
    const HKLM = &H80000002
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    sKeyPath = "SOFTWARE\" & sCompanyName & "\MobileBroadband"
    
    LogEntry("Registry key path is HKLM\" & sKeyPath )
    
    oReg.GetStringValue HKLM, sKeyPath, "SIMID", szSIMID

    oReg.CreateKey HKLM,sKeyPath
    oReg.SetStringValue HKLM, sKeyPath, "ScriptVersion", sVersionNo 
    oReg.SetStringValue HKLM, sKeyPath, "ScriptExecuted", Date & " " & Time
    oReg.SetStringValue HKLM, sKeyPath, "SIMID", SIMID

    LogEntry("Evaluating SIMID Change NEWID: " & SIMID & " - OLDID: " & szSIMID )
    if szSIMID <> SIMID then
      LogEntry("  NEW SIM ID TRIGGERING HARDWARE INVENTORY")
      'Run a SMS Hardware Inventory
      Set cpApplet = CreateObject("CPAPPLET.CPAppletMgr")
      Set actions = cpApplet.GetClientActions
      For Each action In actions
         If Instr(action.Name,"Hardware Inventory") > 0 Then
               action.PerformAction
         End If
      Next
    end if
LogEntry("Ending Script")
wscript.quit


