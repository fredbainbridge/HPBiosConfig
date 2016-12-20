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

SystemSettings.Add "USB Storage Boot", "*Disable"
OSDVariables.Add "USB Storage Boot", "USBBOOT"

SystemSettings.Add "Network (PXE) Boot", "*Enable"
OSDVariables.Add "Network (PXE) Boot", "PXENABLE"

SystemSettings.Add "Legacy Boot Options", "*Enable"
OSDVariables.Add "Legacy Boot Options", "LEGACYBOOTENABLED"

SystemSettings.Add "Legacy Boot Order", "<no legacy boot options available>"
OSDVariables.Add "Legacy Boot Order", "LEGACYBOOTORDER"

SystemSettings.Add "UEFI Boot Options", "*Enable"
OSDVariables.Add "UEFI Boot Options", "UEFIBOOTOPTIONS"

SystemSettings.Add "UEFI Boot Order", "HDD:SATA:1,HDD:USB:1,NETWORK IPV6:EMBEDDED:1,NETWORK IPV4:EMBEDDED"
OSDVariables.Add "UEFI Boot Order", "UEFIBOOTORDER"

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

'the case of the ProBook11g
'The SATA AHCI setting is not available on this model.  Assume it is set correctly.
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objItem in colItems    
    if(objItem.Model = "HP ProBook 11 G2") Then
		env("SATAAHCI") = "1"
	End If
Next


dim OSDVar, SystemConfiguredCorrectly
For Each OSDVar in OSDVariables.Keys
	'debug output
	Wscript.Echo "OSDVar: " & OSDVariables(OSDVar) & " | Value: " & env(OSDVariables(OSDVar))
	If (env(OSDVariables(OSDVar)) = 0) then
		SystemConfiguredCorrectly = false
	End If
Next

If(SystemConfiguredCorrectly) Then
	env(SYSTEMCONFIGURED) = 1
Else
	env(SYSTEMCONFIGURED) = 0
End If

Wscript.Echo "OSDVar: SYSTEMCONFIGURED | Value: " & env(SYSTEMCONFIGURED)
	