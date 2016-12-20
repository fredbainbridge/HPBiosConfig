'tester
On Error Resume Next
set env = CreateObject("Microsoft.SMS.TSEnvironment")
'comment this line out before running in a ConfigMgr Task Sequence
'set env = CreateObject("Scripting.Dictionary")
On Error GoTo 0

Set SystemSettings = CreateObject("Scripting.Dictionary")
Set OSDVariables = CreateObject("Scripting.Dictionary")

' Create the settings to look for.  
' SystemSettings takes the value of the HP Bios setting 
' and a string that is expected in the output of a properly configured system
' run BiosConfigUtility64.exe /getvalue on a reference machine to see output
SystemSettings.Add "Sata Emulation", "*AHCI"
OSDVariables.Add "Sata Emulation", "SATAAHCI"

SystemSettings.Add "TPM Device", "*Available"
OSDVariables.Add "TPM Device", "TPMAVAILABLE"

'Unsure if this is needed? 
SystemSettings.Add "TPM State", "*Enable"
OSDVariables.Add "TPM State", "TPMSTATE"

SystemSettings.Add "Clear TPM", "*No"
OSDVariables.Add "Clear TPM", "CLEARTPM"

SystemSettings.Add "TPM Activation Policy", "*No prompts"
OSDVariables.Add "TPM Activation Policy", "TPMACTIVATIONPOLICY"

SystemSettings.Add "Audio Alerts During Boot", "*Disable"
OSDVariables.Add "Audio Alerts During Boot", "AUDIOALERTS"

SystemSettings.Add "CD-ROM Boot", "*Disable"
OSDVariables.Add "CD-ROM Boot", "CDROMBOOT"

'This is disabled since USB booting is required.  
'Once PXE is enabled this can be uncommented out.
'
'SystemSettings.Add "USB Storage Boot", "*Disable"
'OSDVariables.Add "USB Storage Boot", "USBBOOT"
'

SystemSettings.Add "Network (PXE) Boot", "*Enable"
OSDVariables.Add "Network (PXE) Boot", "PXENABLE"

SystemSettings.Add "Legacy Boot Options", "*Enable"
OSDVariables.Add "Legacy Boot Options", "LEGACYBOOTENABLED"

SystemSettings.Add "Legacy Boot Order", "<no legacy boot options available>"
OSDVariables.Add "Legacy Boot Order", "LEGACYBOOTORDER"

SystemSettings.Add "UEFI Boot Options", "*Enable"
OSDVariables.Add "UEFI Boot Options", "UEFIBOOTOPTIONS"

'Boot orders will vary from model to model.
'More investigation needed.  7/26/2016
'
'SystemSettings.Add "UEFI Boot Order", "HDD:USB:1,NETWORK IPV6:EMBEDDED:1"
'OSDVariables.Add "UEFI Boot Order", "UEFIBOOTORDER"
'

SystemSettings.Add "Wake on LAN on DC mode", "*Enabled"
OSDVariables.Add "Wake on LAN on DC mode", "WAKEONLANDC"

SystemSettings.Add "Configure Legacy Support and Secure Boot", "*Legacy Support Disable"
OSDVariables.Add "Configure Legacy Support and Secure Boot", "SECUREBOOT"

SystemSettings.Add "Wake On LAN", "*Boot to Hard Drive"
OSDVariables.Add "Wake On LAN", "WOL"

SystemSettings.Add "Configure Option ROM Launch Policy", "*All UEFI"
OSDVariables.Add "Configure Option ROM Launch Policy", "ROMPOLICY"

SystemSettings.Add "Hyperthreading", "*Enable"
OSDVariables.Add "Hyperthreading", "HYPERTHREADING"

SystemSettings.Add "Multi-processor", "*Enable"
OSDVariables.Add "Multi-processor", "MULTIPROC"

SystemSettings.Add "Virtualization Technology (VTx)", "*Enable"
OSDVariables.Add "Virtualization Technology (VTx)", "VTX"

SystemSettings.Add "Virtualization Technology for Directed I/O (VTd)", "*Enable"
OSDVariables.Add "Virtualization Technology for Directed I/O (VTd)", "VTD"

Dim Setting
Dim OSDVariableName
For Each Setting in SystemSettings.Keys
	strCommand = "BiosConfigUtility64.exe /getvalue:""" & Setting & """"
	Set oShell = CreateObject("Wscript.Shell")
	Set oExec = oShell.Exec (strCommand)

	ExecOutput = oExec.StdOut.ReadAll
	ExecOutput = Replace (ExecOutput, vbCr,"")
	ExecOutput = Replace (ExecOutput, vbLf,"")
	'Wscript.Echo(ExecOutPut)
	OSDVariableName = OSDVariables(Setting)
	
	if(InStr(ExecOutPut, SystemSettings(Setting) ) <> 0) Then 
		   env(OSDVariableName) = "1"
	Else
		   env(OSDVariableName) = "0"
End If
Next

'the case of the ProBook11g and ZBook Studio G3
'The SATA AHCI setting is not available on these models.  Assume it is set correctly.
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objItem in colItems    
    'This If statement is for all devices that are not able to modify the AHCI settings.
	If(objItem.Model = "HP ProBook 11 G2" Or objItem.Model = "HP ZBook Studio G3" Or objItem.Model = "HP ProOne 400 G2 20-in Touch AiO" Or objItem.Model = "HP Elite x2 1012 G1 " Or objItem.Model = "HP EliteBook 840 G3" Or objItem.Model = "HP Z240 SFF Workstation" Or objItem.Model = "HP EliteDesk 800 G2 DM 35W") Then
		env("SATAAHCI") = "1"
	End If
	If(objItem.Model = "HP ProOne 400 G2 20-in Touch AiO") then 'no battery on this model
		env("WAKEONLANDC") = "1"
		env("HYPERTHREADING") = "1" 'Not an option on this model.
		env("LEGACYBOOTENABLED") = "1" 'Not configurable on this device. UEFI or Legacy.  No hybrid.
	End If
	If(objItem.Model = "HP Elite x2 1012 G1 ") then 'WOL not an option
		env("WAKEONLANDC") = "1"
		env("WOL") = "1"
	End If
	If(objItem.Model = "HP Z240 SFF Workstation" Or objItem.Model = "HP EliteDesk 800 G2 DM 35W") Then
		env("WAKEONLANDC") = "1"
		env("LEGACYBOOTENABLED") = "1"
	End If
Next



dim OSDVar, SystemConfiguredCorrectly
SystemConfiguredCorrectly = true
For Each OSDVar in OSDVariables.Keys
	'debug output
	Wscript.Echo "OSDVar: " & OSDVariables(OSDVar) & " | Value: " & env(OSDVariables(OSDVar))
	If (env(OSDVariables(OSDVar)) = 0) then
		SystemConfiguredCorrectly = false
	End If
Next

If(SystemConfiguredCorrectly) Then
	env("SYSTEMCONFIGURED") = 1
Else
	env("SYSTEMCONFIGURED") = 0
End If

Wscript.Echo "OSDVar: SYSTEMCONFIGURED | Value: " & env("SYSTEMCONFIGURED")
	