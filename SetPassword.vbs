'tester
On Error Resume Next
set env = CreateObject("Microsoft.SMS.TSEnvironment")
'comment this line out before running in a ConfigMgr Task Sequence
'set env = CreateObject("Scripting.Dictionary")
On Error GoTo 0

dim pwAlreadySet
dim correctPWSet
dim PasswordFile1
dim PasswordFile2
'Two files are used to check the existing password. 
'This should be modified per environment
PasswordFile1 = "MPS_SYSTEM_PASSWORD.bin"
PasswordFile2 = "MPS_SYSTEM_PASSWORD2.bin"

strCommand = "BiosConfigUtility64.exe /npwd:""" & PasswordFile1 & """"
Set oShell = CreateObject("Wscript.Shell")
Set oExec = oShell.Exec (strCommand)

ExecOutput = oExec.StdOut.ReadAll
ExecOutput = Replace (ExecOutput, vbCr,"")
ExecOutput = Replace (ExecOutput, vbLf,"")
WScript.echo ExecOutput
'Is a password already set?
if(InStr(ExecOutPut, "Password is set, but no password file is provided" ) <> 0) Then 
        pwAlreadySet = "1"
Else
        if(InStr(ExecOutPut, "Successfully modified Setup Password" ) <> 0) Then
                correctPWSet = "1"
        end if
End If

'If a PW is already set, is it correct?
if(pwAlreadySet = 1) Then
        strCommand = "BiosConfigUtility64.exe /npwd:""" & PasswordFile1 & """ /cpwdfile:""" & PasswordFile2 & """"
        Set oShell = CreateObject("Wscript.Shell")
        Set oExec = oShell.Exec (strCommand)

        ExecOutput = oExec.StdOut.ReadAll
        ExecOutput = Replace (ExecOutput, vbCr,"")
        ExecOutput = Replace (ExecOutput, vbLf,"")
        WScript.echo ExecOutput
        if(InStr(ExecOutPut, "Successfully modified Setup Password" ) <> 0) Then
                correctPWSet = "1"
        else
                correctPWSet = "0"
        end if
end If

env("CORRECTPW") = correctPWSet
WScript.echo "Password Already Set: " & pwAlreadySet 
WScript.echo "Password Correctly Set: " & correctPWSet 

'If there was a previous pw set, fail the task sequence.
if(correctPWSet = "0") then
    On Error Resume Next
    set oTSProgressUI = CreateObject("Microsoft.SMS.TSProgressUI")
    oTSProgressUI.CloseProgressDialog()
    Dim Message 
    Message = "An unexpected system password was found on this device.  Please rectify this in the firmware and retry.  Cannot continue." + vbCr + vbLf
    MsgBox Message & chr(13) & chr(13) & "Press OK to continue.",0, "Warning"
    Wscript.quit 1
end if