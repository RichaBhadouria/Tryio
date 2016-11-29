'==============================================================================================================
'
' NAME: SMS Package: 
'
' AUTHOR: <your name here>. , Bechtel Corporation
' DATE  :   -- Updated - <DATE UPDATED>
'
' Created from script template version: 21.1
'	Functionality Added:
'				1. It is now compatible with Win8.1, Win10, Win 2012 & Win 2012R2
'				3. Check for higher verison with same DisplayName if found quits the script.
'	Bug Fix:
'			1. Fixed reboot issue for Update and component based servicing.
'			2. Fixed issue for Log subroutine giving error on machine having language as Taiwanese
'			3. Check for higher verison corrected
'			4. CheckVersion now works independent of script bitness
'
' COMMENT: 1. This section should be used to describe the package, and any caveats or error codes. 
'             Error Codes: 911 - Installation Failed
'						921 - Installation abandoned, user not an admin
'						901 - Operating System Not Supported
'						900 - Product is already installed
'						950 - User has chosen to postpone.
'						800 - User cancelled the Pending Reboot required.
'						850 - Higher Version installed
'          2. This template has feature to handle two exe (i.e. for same application, depending on bitness).
'         3. Always write generalized code and try to initilize everything at top. 
'        4. Always provide comments describing your change and reason for any deviations.
'==============================================================================================================
Option Explicit

Dim objNetwork, objShell, objFileSys, objEnv, objReg, objWMIService,objItem, colItems, objShellApp, objLocator, objCtx, objReg64, objServices, oReg 
Dim strComputer, strCurrDir, vDecision, strCmd, strRC, strPFpath, strUninstallPath, strOSver, strOSname, intOSType  
Dim blnInstalled, blnChkInstall, bln64BitOS, bln64BitApp, blnXenApp, blnExecute, blnSrv, blnMan, blnElev, blnAdmin 
Dim strAlreadyInst_Msg, strCancel_Msg, strSuccess_Msg, strNoReboot_Msg, strFailed_Msg, strWrongOS_Msg, strNonAdmin_Msg, str_Higher_Msg 
Dim objOpSys, colOpSys, strProdCode, blnFileRenameOperations, blnComponentBasedServicing, blnWindowsUpdateReboot, bln64VersionInstalled
Dim strKeyPath, strValueName, strValue, strSuffix, strUserName, blnSCCMRebootPending, blnSCCMHardRebootPending
Dim objLogFile, strLogName, strLogText, strLogPath
Dim OpSys, OpSysSet, blnInstalled1

Set objNetwork = WScript.CreateObject("WScript.Network")
Set objShellApp = CreateObject("Shell.Application")
Set objShell = WScript.CreateObject("Wscript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objEnv = objShell.Environment("PROCESS")

strComputer = objNetwork.ComputerName

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Const XenAppService = "IMAService"

Const strProdCode64 = "{C7A2BFAF-8AB2-4EFB-AECA-2AE97E2A7F20}" '<name as under HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall can be product code or actual name >
Const strProdCode32 = "{ProductCodeHere}" '<name as under HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall can be product code or actual name >
Const strDispName = "Code42 CrashPlan" '<DisplayName data under strProdCode key>
Const strDispVer = "5.3.0.344" '<DisplayVersion data under strProdCode key>
Const strPkgName = "Code42 CrashPlan" '<Friendly name of the package used in messages boxes>
Const strLogFile = "Code42CrashPlan.log" '<variable part of the log file name, ending in .log - e.g. AcrobatPro_v9.3.log>
bln64BitApp	= true'<identifies the bitness of the app, not a constant so it can be changed at run-time>

Const strProdCodeold = "{393E4C89-67E9-43BF-AD29-94D19F7624F7}"  
Const strDispNameOld = "Connected Backup/PC Agent"
Const strDispVerOld = "8.6" 


strCurrDir = objFileSys.GetParentFolderName(WScript.ScriptFullName)
strLogPath = objEnv("SYSTEMDRIVE") & "\BECNET\SMSLogs\"
strLogName = strLogPath & "SMS-" & strLogFile '<logfile name .. Example - SMS-WinServer2003SP2.log>

Log " " 
Log "Starting log of SMS Package: " & strPkgName

'==========================================================================
' Installation Messages
'==========================================================================
strAlreadyInst_Msg = _
	strPkgName & " is already installed on this machine." & VbCrLf & VbCrLf & _
	"Do you want to reinstall?"

strCancel_Msg = _
	"The installation of " & strPkgName & " has been cancelled." & VbCrLf & VbCrLf & _
	"Click OK to close."

strSuccess_Msg = _
	"The installation of " & strPkgName & " was successful." & VbCrLf & VbCrLf & _
	"Click OK to close."

strNoReboot_Msg = _
	"System restart was deferred. Please restart this system at your convenience." & VbCrLf & vbCrLf & _
	"Click OK to close this program."

strFailed_Msg = _
	"The installation of " & strPkgName & " has failed." & VbCrLf & _
	"Contact your local IS&T helpdesk for support." & VbCrLf & VbCrLf & _
	"Inspect " & strLogPath & "MSI-" & strLogFile & " for details of the problem." & VbCrLf & VbCrLf & _
	"Click OK to close."

strWrongOS_Msg = _
	strPkgName & " can not be installed on this operating system." & VbCrLf & VbCrLf & _
	"Click OK to close this program."

strNonAdmin_Msg = _
	strPkgName & " cannot be installed; user account " & objEnv("USERNAME") & " is not an administrator on " & strComputer & "." & VbCrLf & VbCrLf & _
	"Click OK to close this program."
	
str_Higher_Msg= _
				"Setup has detected that a higher version of " & strPkgName & " is already installed on your system." & VbCrLf & VbCrLf & _
				"If you still want to install this application, please remove the newer version from Add Remove Program and try installation again from Software Center." & VbCrLf & VbCrLf & _
				"Please click on OK to close the program."
	
	
GetNames		' get Computer name and logged in user ID

GetArgs		' Check for command-line arguments

CheckOSVer		' Check State of OS, environment

CheckHigher		' Checks for higher version, Quits program if found


Log "Calling CheckVersion for checking presence of " & strDispNameOld
CheckVersion strProdCodeold,strDispNameOld,strDispVerOld, blnInstalled1		' Check pre-existing installation state

If blnInstalled1 Then
	 Log "Connected Backup/PC Agent found. Proceeding for uninstallation"
	 Log "Calling uninstallConnect subRoutine"
	 uninstallConnect 
Else 
	 Log "Connected Backup/PC Agent not found. Continuing further."
End If



Sub uninstallConnect

  	Log "Starting uninstallConnect sub routine"
	 strCmd= "MsiExec.exe /x " & strProdCodeold & " /passive"
	 ChangeMode "/install"		' change to install mode (XenApp servers only)
	Log "Executing command: " & strCmd
	strRC = objShell.Run(strCmd,0,True)
	
	Log "Calling CheckVersion again to check whether " & strDispNameOld & " has been uninstalled or not."
	CheckVersion strProdCodeold,strDispNameOld,strDispVerOld, blnInstalled1		' Check pre-existing installation state
	If blnInstalled1 Then
		Log "Fail to uninstall Connected Backup/PC Agent. Quitting with error code ."
		WScript.Quit
	Else
 		Log "Connected Backup/PC Agent is uninstalled successfully. continuing further"
	End If
	Log "Ending uninstall sub routine"
End Sub


CheckVersion strProdCode,strDispName,strDispVer, blnInstalled		' Check pre-existing installation state

If Not blnElev Then ' ChkPendingReboot need not run a second time after elevation
	Call ChkPendingReboot		' Check pre-existing reboot state
End If

ChkAdmin		' Check Admin Context

'==========================================================================
' Installation Decisions
'==========================================================================
If blnMan Then
	If blnInstalled Then
		Log "Mandatory argument specified, but " & strPkgName & " is already installed. Wscript.Quitting script with error code: 900"
		Wscript.Quit(900)
	Else		' Mandatory and not installed
		CheckPostpone		' check to see if postponements may be active.  If not, return here to proceed with install.
		DoInstall
	End If
Else		' non-mandatory branch
	If blnInstalled Then
		Log strPkgName & " is installed. Asking user if he/she wants to reinstall."
		vDecision = MsgBox(strAlreadyInst_Msg, vbYesNo + vbQuestion + vbDefaultButton2, strPkgName & " Already Installed")
		If vDecision = vbYes Then
			Log "User answered yes, proceeding to uninstall."
			ReInstall		' handles both uninstall, install
		Else
			Log "User answered no, cancelling installation with a message box.."
			MsgBox strCancel_Msg, 48, "Installation Cancelled"
			Wscript.Quit(0)
		End If
	Else			' main optional install  -- no postpone check for optional
		DoInstall
	End If
End If
Log "End of install decision tree.  Do not expect to reach following exit."
Wscript.Quit(0)		' a normal (but unexpected) exit

Sub ReInstall
	strCmd = "MsiExec.exe /x " & strProdCode & strSuffix & " /norestart /log " & strLogPath & "MSI-Uninstall-" & strLogFile
	ChangeMode "/install"		' change to install mode (XenApp servers only)
	Log "Executing: " & strCmd
	Call objShell.Run(strCmd,0,True)
	Log "Calling DoInstall sub."
	DoInstall
End Sub		' ReInstall

Sub DoInstall
'==========================================================================
' DoInstall Subroutine
'==========================================================================

	Log "Starting DoInstall sub."
	
	objEnv("SEE_MASK_NOZONECHECKS") = 1
	
	If bln64BitApp Then 
		'strCmd = "msiexec /i """ & strCurrDir & "\setup64.msi""" & strSuffix & " /norestart /log " & strLogPath & "MSI-" & strLogFile   'NAME = friendly name of the log file
		 strCmd= """" & strCurrDir & "\code42-package\Install_Code42_CrashPlan.bat"""
	Else 
	'	strCmd = "msiexec /i """ & strCurrDir & "\AgentSetup.msi""" & strSuffix & " /norestart /log " & strLogPath & "MSI-" & strLogFile   'NAME = friendly name of the log file
	End If 
	
	ChangeMode "/install"		' change to install mode (XenApp servers only)
	Log "Executing command: " & strCmd
	strRC = objShell.Run(strCmd,0,True)
		
	Log "Calling CheckVersion sub."
	CheckVersion strProdCode,strDispName,strDispVer,blnInstalled
	If blnInstalled Then
		PostInstall		' place any post-installation steps in the subroutine at the bottom
		ChangeMode "/execute"		' change to execute mode (XenApp servers only)
		If blnMan Then
			Log strPkgName & " installed successfully. Exiting."
			Wscript.Quit(0)
		Else
			Log strPkgName & " installed successfully. Informing user, Wscript.Quitting script without error"
			MsgBox strSuccess_Msg, vbOKOnly + vbInformation, "Installation Successful"
		End If 
	Else
		ChangeMode "/execute"		' change to execute mode (XenApp servers only)
		If blnMan Then
			Log strPkgName & " failed to install, informing user, Wscript.Quitting with error code 911."
			Log "Inspect " & strLogPath & "MSI-" & strLogFile & " for details."
			Wscript.Quit(911)
		Else
			Log strPkgName & " failed to install, informing user, Wscript.Quitting with error code 911."
			MsgBox strFailed_Msg, 16, "Installation Failed"
			Wscript.Quit(911)
		End If 
	End If
	
	objEnv.Remove("SEE_MASK_NOZONECHECKS")

	Log "Ending DoInstall sub."

End Sub		' DoInstall

Sub PostInstall
		'<place any post install pieces here>
End Sub		' PostInstall

'========================================================================================
'-----STANDARD SUBROUTINES i.e. which are unchanged---- 
'-----If you change any of the functions below, please bring it above this section----
'========================================================================================

Sub DoReboot
'==========================================================================
' Reboot Section  - in case required
'==========================================================================

	Log "Starting DoReboot subroutine."
	
	Log "Executing Reboot. Logging complete at " & Now	   
	Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
	For Each OpSys In OpSysSet
		OpSys.Reboot()
	Next

End Sub

Sub ChkPendingReboot 
'---------------------------Method 1--------------------------------------
	Dim arrValues, strReturn, strValue, strFileOps, strKeyPath, strValueName
	
	Log "Inside ChkPendingReboot subroutine." 
	
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
	strValueName = "PendingFileRenameOperations"
	
	strReturn =objReg.GetMultiStringValue( HKEY_LOCAL_MACHINE,strKeyPath,strValueName,arrValues)
	
	If (strReturn = 0) And (Err.Number = 0) Then  
	    blnFileRenameOperations = True
		For Each strValue In arrValues
			strFileOps = strFileOps & chr(13) & strValue
		Next
	Else
		If Err.Number = 0 Then
			Log "No Pending File Operations Found"
		Else
			Log "Check Pending File Operations failed. Error = " & Err.Number
		End If
	End If 
	
	'---------------------------Method 2--------------------------------------
    Dim strKeyPath1, hDefKey1, arrSubKeys
    hDefKey1 = HKEY_LOCAL_MACHINE
    strKeyPath1 = "Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
	If bln64BitOS Then
		If objReg64.EnumKey(hDefKey1, strKeyPath1, arrSubKeys) = 0 Then
	        blnComponentBasedServicing = True
	        Log "Component Based Servicing Pending Reboot Operations found. "
	    Else
	     Log "No Component Based Servicing Pending Reboot Operations found. "
	    End If
	Else
	    If objReg.EnumKey(hDefKey1, strKeyPath1, arrSubKeys) = 0 Then
	        blnComponentBasedServicing = True
	        Log "Component Based Servicing Pending Reboot Operations found. "
	    Else
	     Log "No Component Based Servicing Pending Reboot Operations found. "
	    End If
    End If
	
	'---------------------------Method 3------------------------------------------
    Dim strKeyPath2, hDefKey2, arrSubKeys2
    hDefKey2 = HKEY_LOCAL_MACHINE
    strKeyPath2 = "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
	If bln64BitOS Then
		If objReg64.EnumKey(hDefKey2, strKeyPath2, arrSubKeys2) = 0 Then
	            blnWindowsUpdateReboot = True
	            Log "Windows Update Pending Reboot Operations found "
	   Else
	          Log "No Windows Update Pending Reboot Operations found "         
	   End If
	Else
       If objReg.EnumKey(hDefKey2, strKeyPath2, arrSubKeys2) = 0 Then
	            blnWindowsUpdateReboot = True
	            Log "Windows Update Pending Reboot Operations found "
	   Else
	          Log "No Windows Update Pending Reboot Operations found "         
	   End If
	End If
	   
	If blnFileRenameOperations Or blnComponentBasedServicing Or blnWindowsUpdateReboot Then
		If blnMan = True Then		
			Log "Pending Reboot Operations Found: " & strFileOps & ". But blnMan = True. Setup will continue"
		Else
			Log "Pending Reboot Operations Found: " & strFileOps & ". Ask user if setup should continue"
			vDecision = MsgBox("Setup has detected that this computer has a pending reboot." & vbCr & "You can continue with the installation but it is recommended that you reboot the computer and run setup again." &vbCr & vbCr & "Click Ok to continue", vbInformation+vbOKcancel,"Pending reboot")
			If vDecision = vbOK Then
				Log "User answered yes, setup will continue..."
			Else
				Log "User answered no, cancelling installation with a message box.."
				MsgBox strCancel_Msg, 48, "Installation Cancelled"
				Wscript.Quit(800)
			End If		  
		End If 
	End If
	Log "Ending ChkPendingReboot subroutine." 

End Sub		' ChkPendingReboot

Sub CheckOSVer
'==========================================================================
' CheckOSVer Subroutine
'==========================================================================

	Log "Starting CheckOSVer subroutine."
	
	If Right(objEnv("PROCESSOR_ARCHITECTURE"),2) = "64" Or Right(objEnv("PROCESSOR_ARCHITEW6432"),2) = "64" Then
		Log "Operating system check detected Windows 64-bit"
		bln64BitOS = True
		If bln64BitApp Then
			strPFpath = objEnv("SYSTEMDRIVE") & "\Program Files\" ' objEnv("SYSTEMDRIVE") = C:
			strProdCode = strProdCode64
		Else
			strPFpath = objEnv("SYSTEMDRIVE") & "\Program Files (x86)\"
			strProdCode = strProdCode32
		End If
		
		Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
		objCtx.Add "__ProviderArchitecture", 64
		objCtx.Add "__RequiredArchitecture", True
		Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
		Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx)
		Set objReg64 = objServices.Get("StdRegProv") 
	Else
		Log "Operating system check detected Windows 32-bit"
		strPFpath = objEnv("SYSTEMDRIVE") & "\Program Files\"
		strProdCode = strProdCode32
		bln64BitOS = False
		bln64BitApp = False
	End If
	
	strUninstallPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

	Set colOpSys = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each objOpSys In colOpSys

		strOSver = objOpSys.Version				' FOR WIn7x64 enterprise, Version = "6.1.7601"
		strOSname = objOpSys.Caption			' FOR WIn7x64 enterprise, Caption = "Microsoft Windows 7 Enterprise "
		intOSType = objOpSys.ProductType		' 1=workstation, 2=DC, 3=server
		If intOSType = 1 Then
			blnSrv = False
		Else
			blnSrv = True
		End If
		
		If InStr(UCASE(objOpSys.Caption), "SERVER 2003") <> 0 Then  
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion
		
		ElseIf InStr(UCASE(objOpSys.Caption), "SERVER® 2008") <> 0 Then  
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion 
			If Not blnElev Then Elevate
		
		ElseIf InStr(UCASE(objOpSys.Caption), "SERVER 2008 R2") <> 0 Then  
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion    
			If Not blnElev Then Elevate
			
		ElseIf InStr(UCASE(objOpSys.Caption), "SERVER 2012 R2") <> 0 Then  
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion    		
			If Not blnElev Then Elevate		

		ElseIf InStr(UCASE(objOpSys.Caption), "SERVER 2012") <> 0 And InStr(UCASE(objOpSys.Caption), "SERVER 2012 R2") = 0 Then  
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion    
			If Not blnElev Then Elevate
		
		ElseIf InStr(UCASE(objOpSys.Caption), "SERVER 2016") <> 0 Then  
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion    
			If Not blnElev Then Elevate
			
		ElseIf InStr(UCASE(objOpSys.Caption), "WINDOWS 7") <> 0 Then 'searches for "WINDOWS 7" and returns the position no as 11 in "Microsoft Windows 7 Enterprise"
			Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion
			If Not blnElev Then Elevate
		
       ElseIf InStr(UCASE(objOpSys.Caption), "WINDOWS 8.1") <> 0 Then
          Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion
          If Not blnElev Then Elevate
      
       ElseIf InStr(UCASE(objOpSys.Caption), "WINDOWS 8") <> 0 And InStr(UCASE(objOpSys.Caption), "8.1") = 0 Then
         Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion
         If Not blnElev Then Elevate
      
      ElseIf InStr(UCASE(objOpSys.Caption), "WINDOWS 10") <> 0 Then 'searches for "WINDOWS 10" and returns the position no as 11 in "Microsoft Windows 10 Enterprise"
         Log "Operating system check detected: " & objOpSys.Caption & " Service pack: " & objOpSys.ServicePackMajorVersion
         If Not blnElev Then Elevate		
	  Else
			Log "This Operating System is not applicable: " & objOpSys.Caption
			If Not blnMan Then
				MsgBox strWrongOS_Msg, 16, "Operating System NOT Compatible"
			End If
			Wscript.Quit(901)
		End If
	Next

	' Determine if this is a Citrix XenApp server
	If ServicePresent(XenAppService) Then
		blnXenApp = True
	Else
		blnXenApp = False
	End If 

End Sub		' CheckOSVer

Sub CheckHigher
	
	Log "Starting CheckHigher Subroutine"
	
	' Set reg object based on type of applicaiton
	If bln64BitOS Then
		Set oReg = objReg64
		Log "Setting Reg object to 64 bit. Enumrating 64 Bit hive for higher version"
		EnumHives
	End If
	
	Set oReg = objReg
	Log "Setting Reg object to 32 bit. Enumrating 32 Bit hive for higher version"
	EnumHives
	
	Log "Ending CheckHigher subroutine"

End Sub																		' End CheckHigher

Sub EnumHives
	On Error Resume Next

	Dim strValueVersion, arrSubKeys, strSubkey, strSubKeyPath, arrValueNames, arrTypes, strValueName, strValue, i
	Log "Starting function EnumHives"

	oReg.EnumKey HKEY_LOCAL_MACHINE, strUninstallPath, arrSubKeys														' Getting all uninstall hive
	
	For Each strSubkey In arrSubKeys																					' Processing all hives
		strSubKeyPath = strUninstallPath & strSubkey
	 	
	 	oReg.EnumValues HKEY_LOCAL_MACHINE, strSubKeyPath, arrValueNames, arrTypes
		
		For i = LBound(arrValueNames) To UBound(arrValueNames)
	    	strValueName = arrValueNames(i)
			If (arrTypes(i) = REG_SZ) Then																		' Consindring only Reg_sz entries
	    	    oReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKeyPath, strValueName, strValue
	        	If (strValueName = "DisplayName") Then
			        If (strValue = strDispName) Then															' Application with same display name found
			        	Log "Application with same name display name found. Checking for version."
			        	oReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKeyPath, "DisplayVersion", strValueVersion		' Getting  Display Version
			        	Log "Display version of application found: " & strValueVersion
			        	If CCompare (strValueVersion, strDispVer) Then															' Comparing the versions
			        		If blnMan Then
			        			Log "Higher version: " & strValueVersion & " of application found installed. Quiting script with exiting script with error code 850."
			        		Else
			        			Log "Higher version: " & strValueVersion & " of application found installed. Informing user and quiting script with exiting script with error code 850."
			        			MsgBox str_Higher_Msg, vbOKOnly+vbInformation, "Newer version Installed"
			        		End If
			        		WScript.Quit (850)
			        	Else
			        		Log "Lower or Same version: " & strValueVersion & " of application found installed. Proceeding with installation"
			        	End If
			       End If
	        	End If
	    	End If
		Next
	Next
	
	Log "Ending function EnumHives"

End Sub	'EnumHives


Sub CheckVersion(strProdCode,strDispName,strDispVer, blnChkInstall)
'==========================================================================
' CheckVersion Subroutine
'==========================================================================
' NB: 4th parameter permits handling multiple product checks independently.

	Dim Inparams, Outparams, blnChkInstall64, blnChkInstall32
	Log "Starting CheckVersion subroutine."

' In case CheckVersion is used after an installation that bounces WMI, reset the WMI objects...  
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	strKeyPath = strUninstallPath & strProdCode
	strValueName = "DisplayVersion"

	
	If bln64BitOS Then		' Use ExecMethod to call the GetStringValue method
		Log "Checking for 64 Bit application presence at: " & strKeyPath & "\" & strValueName
		Set Inparams = objReg64.Methods_("GetStringValue").Inparameters
		Inparams.Hdefkey = HKEY_LOCAL_MACHINE
		Inparams.Ssubkeyname = strKeyPath
		Inparams.Svaluename = strValueName
		Set Outparams = objReg64.ExecMethod_("GetStringValue", Inparams,,objCtx)
		If Outparams.sValue = strDispVer Then
			blnChkInstall64 = True
			bln64VersionInstalled = True
			Log "64 Bit version of " & strDispName & " is installed on this system. Setting Boolean value: blnInstalled = True"
		Else
			blnChkInstall64 = False
		End If
	End If
	
	Log "Checking for 32 Bit application presence at: " & strKeyPath & "\" & strValueName
	objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	If strValue = strDispVer Then
		blnChkInstall32 = True
		Log "32 Bit version of " & strDispName & " is installed on this system. Setting Boolean value: blnInstalled = True"
	Else
		blnChkInstall32 = False
	End If 	
	
	If (blnChkInstall64 Or blnChkInstall32) Then
		blnChkInstall = True
	Else 
		blnChkInstall = False
	End IF
	
	If Not blnChkInstall Then
		Log strDispName & " is NOT installed on this system. Setting Boolean value: blnInstalled = False"
	End If		  
	
	Log "Ending CheckVersion subroutine."

End Sub		' CheckVersion

Sub GetNames
' log computer, user names
	Set colItems = objWMIService.ExecQuery("Select * From Win32_ComputerSystem")
	For Each objItem in colItems
		strUserName = objItem.UserName
		Log "Computer name: " & objNetwork.ComputerName & ".  Current user logged in: " &  strUserName
	Next
End Sub		' GetNames

Sub GetArgs
'==========================================================================
' Handle Arguments: 
'==========================================================================
	Dim i	
	Log "Setting DEFAULT argument variables: blnMan = False"
	blnMan = False
	blnElev = False
	
	Log "Checking for arguments."
	
	If Wscript.arguments.Count < 1 Then
		Log "Script has been executed without arguments. Continuing script."   
	ElseIf Wscript.arguments.Count > 2 Then
		Log "Too many arguments defined...will ignore. Continuing script."   
	Else
		For i = 0 To WScript.Arguments.Count-1
			Select Case UCase(WScript.Arguments.Item(i))
				Case "MANDATORY"
					blnMan = True
					Log "Mandatory argument specified, assigning variable blnMan = True"
				Case "ELEVATE"	
					blnElev = True	'	Signals this execution is already elevated, don't re-elevate.
					Log "Restarting script in elevated mode."
				Case Else
					Log "Argument " & WScript.Arguments.Item(i) & " not recognized, ignored."
			End Select
		Next	
	End If

	If blnMan = True Then
		strSuffix = " /quiet"
	Else
		strSuffix = " /passive"
	End If
End Sub		' GetArgs

Sub Log(strLogText)
'==========================================================================
' Logging Subroutine
'==========================================================================
	SetLocale (1033)																	' Setting language for writing log Default language of script not changed as it may effect functionality of application
	If not objFileSys.FolderExists (objEnv("SYSTEMDRIVE") & "\BECNET") Then
		 objFileSys.CreateFolder objEnv("SYSTEMDRIVE") & "\BECNET\"
	End If  
	
	If not objFileSys.FolderExists (objEnv("SYSTEMDRIVE") & "\BECNET\SMSLogs") Then
		 objFileSys.CreateFolder objEnv("SYSTEMDRIVE") & "\BECNET\SMSLogs\"
	End If   
	
	If Not objFileSys.FileExists(strLogName) Then
		Set objLogFile = objFileSys.CreateTextFile(strLogName,8)
		objLogFile.WriteLine "Log File: " & strLogName & " created " & Now
		objLogFile.WriteLine " "
		objLogFile.Close
	End If 
		
	Set objLogFile = objFileSys.OpenTextFile(strLogName,8)
	objLogFile.WriteLine Now & ": " & strLogText
	objLogFile.Close
	SetLocale (0)
	
End Sub		' Log

Sub Elevate
'==========================================================================
' Subroutine used on Vista or later to restart script with UAC elevation
'==========================================================================
	Dim strEnvUserName		: strEnvUserName = UCase(objEnv("USERNAME")) 
	Dim i 
	If strEnvUserName = "SYSTEM" Or strEnvUserName = strComputer & "$" Then		
		blnElev = True
		Log "Program running under computer account, elevation not required."
	Else		
		Dim strElev
		strElev = Chr(34) & strCurrDir & "\" & WScript.ScriptName & """ ELEVATE"
		For i = 0 To WScript.Arguments.Count-1 
			strElev = strElev & " " & WScript.Arguments.Item(i)
		Next
		Log "Restarting script using: " & strElev
		objShellApp.ShellExecute "wscript.exe", strElev, "", "runas", 1
		Wscript.Quit(0)
	End If
End Sub		' Elevate

Sub ChkAdmin
'==========================================================================
' Subroutine to check if user is an administrator
'==========================================================================
	Dim strAdmTestKey 	: strAdmTestKey = "System\CurrentControlSet\Control"
	Dim strEnvUserName	: strEnvUserName = objEnv("USERNAME")
	Dim strTestValue		: strTestValue = "nullval"
	Dim strAdmTestVal		: strAdmTestVal = "becadmtest"
	blnAdmin = False
	If strEnvUserName = "SYSTEM" Or strEnvUserName = strComputer & "$" Then
		blnAdmin = True
		Log "Program running under " & strEnvUserName & " presumed to have admin rights, no check made."
	Else
		Log "Checking admin rights for: " & strEnvUserName
		On Error Resume Next
		Err.Clear
		objReg.SetStringValue HKEY_LOCAL_MACHINE,strAdmTestKey,strAdmTestVal,strEnvUserName       'creating key becadmtest and setting its value as <username>
		If Err.Number Then
			Log "User " & strEnvUserName & " is not an admin, error code " & Err.Number & "."
		Else
			Err.Clear
			objReg.GetStringValue HKEY_LOCAL_MACHINE,strAdmTestKey,strAdmTestVal,strTestValue     'reading the value of key becadmtest
			If Err.Number Or strTestValue <> strEnvUserName Or IsNull(strTestValue) Then
				Log "User " & strEnvUserName & " is not an admin, error code " & Err.Number & "."
			Else
				blnAdmin = True
				Log "User " & strTestValue & " is an admin."
				objReg.DeleteValue HKEY_LOCAL_MACHINE,strAdmTestKey,"becadmtest"				   'deleting the value becadmtest
			End If
		End If
		On Error GoTo 0
	End If

	If Not blnAdmin Then
		If blnMan Then
			Log "Mandatory installation cannot continue, user not at admin.  Exit 921."
		Else
			Log "Reinstall requested but cannot continue; user is not an admin.  Exit 921."
			MsgBox strNonAdmin_Msg, 16, "Admin Rights Required"
		End If 
		Wscript.Quit(921)
	End If
End Sub			' ChkAdmin

Function ServicePresent(Service)
'==========================================================================
' SMS_SITE_COMPONENT_MANAGER can take longer to stop than msiexec is willing to wait.
'  Here, we stop it first, waiting until it completes.
'==========================================================================

	Dim colServices, objService
	ServicePresent = False
	Set colServices = objWMIService.ExecQuery ("SELECT * FROM win32_Service WHERE Name = '" & Service & "'")
	For Each objService in colServices
		ServicePresent = True
	Next
End Function		' ServicePresent

Sub ChangeMode(UserMode)	
' If this is a XenApp server, change mode to UserMode 	
	If blnXenApp Then	' if not XenApp we don't need to do this
		Dim strCMcmd 	: strCMcmd = "change user " & UserMode
		Select Case UCase(UserMode)
			Case "/INSTALL"
				If blnExecute Then
					Log "Changing mode of XenApp server via command: " & strCMcmd
					Call objShell.Run(strCMcmd,,True) 
					blnExecute = False
				End If
			Case "/EXECUTE"
				If Not blnExecute Then
					Log "Changing mode of XenApp server via command: " & strCMcmd
					Call objShell.Run(strCMcmd,,True) 
					blnExecute = True
				End If
			Case Else
				Log "User mode " & UserMode & " not recognized, ignored."
		End Select
	End If
End Sub		' ChangeMode

Function CCompare (Var1, Var2)
 
 	On Error Resume Next 
 	Dim A, B, C, D, N, i
 	CCompare = False										' Setting Default Value for function
 	
 	A = Split(Var1,".")										' Getting each decimal number in Array
 	B = Split(Var2,".")
 	C= UBound (A)											' Getting lenth of Var 1
 	D= UBound (B)											' Getting lenth of Var 2
 	
 	If C < D Then											' If lenth of variable different choose the lower one
 		N= C
 	Else
 		N= D
 	End If
 	
 	For i= 0 To N											' Compare if difrrence is found exit 
 		If A(i) > B(i) Then
 			CCompare = True
	 		Exit Function
	 	ElseIf A(i) < B(i) Then
	 		Exit Function
	 	End If
 	Next
 	
 	If C > D Then											' If all number are equal then compare lenth
 		CCompare = True
 	End If
 	
 	'If all are equal function return default i.e. CCompare= False
 	
End Function

Sub CheckPostpone ' ---PLEASE READ ALL COMMENTS CAREFULLY----
'==========================================================================
' Postponement Handling.  To activate:
'		(a) Set intPostponMax to a non-zero value
'		(b) Review the text of messages (below).  This needs change as per requirement.
'		(c) Assumes postponements are not allowed (or do not apply) on a server.
'		(d) Assumes running as MANDATORY (optional installs shouldn't need postponements).
'==========================================================================

	Const intPostponeMax = 0				' Max number of postponements allowed for this pkg.  If set = 0, postponements are disabled. If you need 5 postponemets, please initialize 4 here.
	If intPostponeMax = 0 Then Exit Sub		' Bypass postponement handling/checking.

	Dim intPostponeRemain, blnAllowPostpone, strUpgrade_Msg_MAN, strNoPostpone_Msg, dwValue, newdwValue
	Const strPostponeCtKey = "SOFTWARE\Bechtel\SMSPkg"
	Const strPostponeCtVal = ""	' Registry value (name) where COUNT of postponements remaining is kept, e.g. 'Office2010'. Need to be initialized.
	
	Log "Starting Postponement subroutine."
	If blnSrv Then
		blnAllowPostpone = False
		Log "This is a server OS, postponements are not an option."
	Else
		Log "Checking the registry: HKLM\" & strPostponeCtKey & "\" & strPostponeCtVal
		objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strPostponeCtKey,strPostponeCtVal,dwValue
		Log "Registry check returned dwvalue: " & dwValue
		If dwValue = intPostponeMax Then
			blnAllowPostpone = False
			Log "The number of postponements is = to the allowed number. Setting variable: blnAllowPostpone = False."
		ElseIf dwValue < intPostponeMax Then
			blnAllowPostpone = True
			newdwValue = (dwValue + 1)
			objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strPostponeCtKey,strPostponeCtVal,newdwValue
			intPostponeRemain = (intPostponeMax - newdwValue) + 1  
			Log "The number of postponements taken is less than the number allowed. Setting variable: blnAllowPostpone = True."
			Log "Decrementing the number of postponements: intPostponeRemain = " & intPostponeRemain
		Else		' the value name doesn't even exist yet, create it
			blnAllowPostpone = True
			Log "The number of postponements has not been defined at all. Setting variable: blnAllowPostpone = True."
			objReg.CreateKey HKEY_LOCAL_MACHINE,strPostponeCtKey
			dwValue = 0
			objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strPostponeCtKey,strPostponeCtVal,dwValue
			Log "Creating registry key: HKLM\" & strPostponeCtKey & "\" & strPostponeCtVal & "\" & dwValue
			intPostponeRemain = (intPostponeMax - dwValue) + 1
			Log "Specifying variable for remaining postponements: intPostponeRemain = " & intPostponeRemain
		End If
	End If
	
'	Messaging handling ... review before activating.  Text sample below was used in a MS Office script.
	strUpgrade_Msg_MAN = _
		"Message from Bechtel IS&T:" & VbCrLf & VbCrLf & _
		"Bechtel IS&T is proactively upgrading your system to install " & strPkgName & " to help you take advantage of its new features. We appreciate your patience as your system is upgraded." & VbCrLf & VbCrLf & _
		"Once the version upgrade is complete, you will be prompted to restart your machine. Failing to restart may leave Microsoft Office in an unstable state." & VbCrLf & VbCrLf & _
		"Click OK to continue with the upgrade, or Cancel to postpone the installation until a later time." & VbCrLf & _
		"You have " & intPostponeRemain & " postponements remaining, after which the installation will become mandatory."
	
	strNoPostpone_Msg = _
		"Message from Bechtel IS&T:" & VbCrLf & VbCrLf & _
		"Bechtel IS&T is proactively upgrading your system to install " & strPkgName & " to help you take advantage of its new features. We appreciate your patience as your system is upgraded." & VbCrLf & VbCrLf & _
		"There are no further postponements allowed." & VbCrLf & VbCrLf & _
		"Click OK to begin the upgrade process. You will be prompted to restart your system once complete." & VbCrLf & _
		"Note: Failing to restart may leave Microsoft Office in an unusable state."
		
' --- SEE MORE SAMPLEs of messages BELOW. Please fill approved text in these messages---
		
	' strUpgrade_Msg_MAN = _
' 		"Your installation of Citrix receiver must be upgraded to version 14.3.100.10 to meet the security requirements" & VbCrLf & VbCrLf & _
' 		"During the upgrade all applications accessed through the receiver will be unavailable. Please save your work and close any open applications." & VbCrLf & VbCrLf & _
' 		"After the installation a reboot is required for the single sign on feature to function correctly, however reboot is not enforced, please reboot at your earliest convenience. " & vbCrLf & vbCrLf & _
' 		"Click “OK” to proceed.  Click on “Cancel” to defer the upgrade. You have " & intPostponeRemain & " deferrals left." 
' 		 
' 	strNoPostpone_Msg = _
' 		"Your installation of Citrix receiver must be upgraded to version 14.3.100.10 to meet the security requirements" & VbCrLf & VbCrLf & _
' 		"During the upgrade all applications accessed through the receiver will be unavailable. Please save your work and close any open applications." & VbCrLf & VbCrLf & _
' 		"After the installation a reboot is required for the single sign on feature to function correctly, however reboot is not enforced, please reboot at your earliest convenience. " & vbCrLf & vbCrLf & _
' 		"The allowed number of deferrals for this upgrade are exhausted, the upgrade will begin now, Click on OK to proceed." 

	If blnAllowPostpone Then		' postponement is allowed, ask user
		Log "Postponement allowed. Showing mandatory upgrade message to user."
		vDecision = MsgBox(strUpgrade_Msg_MAN, vbOKCancel + vbQuestion + vbDefaultButton2, "Install " & strPkgName & "?")
		If vDecision = vbOK Then
			Log "User chose to upgrade."
		Else
			Log "User chose to postpone. Wscript.Quitting script, exit 950."
			MsgBox strCancel_Msg, 48, "Installation Cancelled"
			Wscript.Quit(950)
		End If 
	Else		' postponement not allowed / used up
		If not blnSrv Then 	' if it's a workstation, inform the user
			MsgBox strNoPostpone_Msg,vbInformation+vbOKOnly,"Installing " & strPkgName
		End If
	End If

Log "Ending CheckPostpone subroutine."

End Sub		' CheckPostpone
