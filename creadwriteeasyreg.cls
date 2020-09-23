VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReadWriteEasyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Registry API's to use


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long


Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpKeyName As String) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function ExpandEnvironmentStrings Lib "advapi32.dll" (lpSrc As String, lpDst As String, ByVal nSize As Long) As Long


Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Variant
    bInheritHandle As Long
    End Type
    'Enum's for the OpenRegistry function


Public Enum HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum


'Enum's for the DataTypes


Public Enum lDataType
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

'Right's for the OpenRegistry
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = ( _
KEY_READ _
And _
KEY_WRITE _
)
Private Const KEY_ALL_ACCESS = ( _
( _
STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or _
KEY_SET_VALUE Or _
KEY_CREATE_SUB_KEY Or _
KEY_ENUMERATE_SUB_KEYS Or _
KEY_NOTIFY Or _
KEY_CREATE_LINK _
) _
And _
( _
Not SYNCHRONIZE _
) _
)
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1


Dim RootHKey As HKeys
Dim SubDir As String
Dim hKey As Long
Dim OpenRegOk As Boolean


Function GetAllSubDirectories() As Variant


    On Error GoTo handelgetdirvalues
    Dim SubKey_Num As Integer
    Dim SubKey_Name As String
    Dim length As Long
    Dim ReturnArray() As Variant
    
    If Not OpenRegOk Then Exit Function
    'Get the Dir List


SubKey_Num = 0



    Do
        length = 256


SubKey_Name = Space$(length)



    If RegEnumKey(hKey, SubKey_Num, SubKey_Name, length) <> 0 Then
        Exit Do
    End If



SubKey_Name = Left$(SubKey_Name, InStr(SubKey_Name, Chr$(0)) - 1)

    ReDim Preserve ReturnArray(SubKey_Num) As Variant
    ReturnArray(SubKey_Num) = SubKey_Name


SubKey_Num = SubKey_Num + 1

Loop

GetAllSubDirectories = ReturnArray
Exit Function
handelgetdirvalues:
GetAllSubDirectories = Null
Exit Function
End Function

'This function will return a true or false when it creates a key


'     for you
'eg.
' Dim MyReg As New CReadWriteEasyReg
' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microso
'     ft") Then
'MsgBox "Couldn't open the registry"
'Exit Sub
' End If
' if MyReg.CreateDirectory("TestDir") then
'Msgbox "Key created"
' else
'msgbox "Couldn't Create key"
' end if
' MyReg.CloseRegistry


Public Function CreateDirectory(ByVal sNewDirName As String) As Boolean

    Dim hNewKey As Long, lpdwDisposition As Long
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    Dim lReturn As Long
    
    If Not OpenRegOk Then Exit Function
    
    lReturn = RegCreateKeyEx(hKey, sNewDirName, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, lpdwDisposition)


    If lReturn = 0 Then
        CreateDirectory = True
    Else
        CreateDirectory = False
    End If

End Function

'This function will return a true or false when it deletes a key


'     for you
    'eg.
    ' Dim MyReg As New CReadWriteEasyReg
    ' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microso
    '     ft") Then
    'MsgBox "Couldn't open the registry"
    'Exit Sub
    ' End If
    ' if MyReg.DeleteDirectory("MyTestDir") then
    'Msgbox "Key Deleted"
    ' else
    'msgbox "Couldn't Delete key"
    ' end if
    ' MyReg.CloseRegistry


Public Function DeleteDirectory(ByVal sKeyName As String) As Boolean

    Dim lReturn As Long
    
    If Not OpenRegOk Then Exit Function
    
    lReturn = RegDeleteKey(hKey, sKeyName)


    If lReturn = 0 Then
        DeleteDirectory = True
    Else
        DeleteDirectory = False
    End If

End Function

'This function will return a array of variant with all the value
'     names in a key
'eg.
' Dim MyVariant As Variant, MyReg As New CReadWriteEasyReg, i As
'     Integer
' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "HardWare\Descrip
'     tion\System\CentralProcessor\0") Then
'MsgBox "Couldn't open the registry"
'Exit Sub
' End If
' MyVariant = MyReg.GetAllValues
' For i = LBound(MyVariant) To UBound(MyVariant)
'Debug.Print MyVariant(i)
' Next i
' MyReg.CloseRegistry


Function GetAllValues() As Variant

    On Error GoTo handelgetdirvalues
    Dim lpData As String, KeyType As Long
    Dim BufferLengh As Long, vname As String, vnamel As Long
    Dim ReturnArray() As Variant, Index As Integer
    
    If Not OpenRegOk Then Exit Function
    
    'Get the Values List
    Index = 0


    Do
        lpData = String(250, " ")
        BufferLengh = 240
        vname = String(250, " ")
        vnamel = 240


        If RegEnumValue(ByVal hKey, ByVal Index, vname, vnamel, 0, KeyType, lpData, BufferLengh) <> 0 Then
            Exit Do
        End If

        vname = Left$(vname, InStr(vname, Chr$(0)) - 1)
        ReDim Preserve ReturnArray(Index) As Variant
        ReturnArray(Index) = vname
        Index = Index + 1
    Loop

    GetAllValues = ReturnArray
    Exit Function
handelgetdirvalues:
    GetAllValues = Null
    Exit Function
End Function

'This function will return a true or false when it creates a valu
'     e for you
'eg.
' Dim MyReg As New CReadWriteEasyReg
' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microso
'     ft") Then
'MsgBox "Couldn't open the registry"
'Exit Sub
' End If
' if MyReg.CreateValue("ValName", "This is written as the value",
'     REG_SZ) then
'Msgbox "Value created"
' else
'msgbox "Couldn't Create Value"
' end if
' MyReg.CloseRegistry


Public Function CreateValue(ByVal sValueName As String, ByVal vWriteThis As Variant, ldValueDataType As lDataType, Optional Multi_SZ_AddtlStrings As Variant) As Boolean

    Dim lpData As String 'The pointer to the value written to the Registry key's value
    Dim cbData As Long 'The size of the data written to the Registry key's value, including termination characters If applicable
    Dim lReturn As Long 'The Error value returned by the Registry Function
    Dim Str As Variant
    
    If Not OpenRegOk Then Exit Function
    


    Select Case ldValueDataType
        Case REG_SZ, REG_EXPAND_SZ
        lpData = vWriteThis & Chr(0)
        cbData = Len(lpData)
        lReturn = RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData)


        If lReturn = 0 Then
            CreateValue = True
        Else
            CreateValue = False
        End If

        Case REG_MULTI_SZ
        lpData = vWriteThis & Chr(0)


        If Not IsMissing(Multi_SZ_AddtlStrings) Then


            If IsArray(Multi_SZ_AddtlStrings) Then


                For Each Str In Multi_SZ_AddtlStrings


                    If Str <> "" And Str <> Chr(0) And Not IsNull(Str) Then
                        lpData = lpData & Str & Chr(0)
                    End If

                Next Str

            Else


                If Multi_SZ_AddtlStrings <> "" And Multi_SZ_AddtlStrings <> Chr(0) And Not IsNull(Multi_SZ_AddtlStrings) Then
                    lpData = lpData & Multi_SZ_AddtlStrings & Chr(0)
                End If

            End If

        End If

        lpData = lpData & Chr(0)
        cbData = Len(lpData)
        lReturn = RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData)


        If lReturn = 0 Then
            CreateValue = True
        Else
            CreateValue = False
        End If

        Case REG_DWORD
        lpData = CLng(vWriteThis)
        cbData = 4
        lReturn = RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData)


        If lReturn = 0 Then
            CreateValue = True
        Else
            CreateValue = False
        End If

        Case Else
        MsgBox "Unable to process that Type of data."
        CreateValue = False
    End Select

End Function

'This function will return a true or false when it deletes a valu
'     e for you
'eg.
' Dim MyReg As New CReadWriteEasyReg
' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microso
'     ft") Then
'MsgBox "Couldn't open the registry"
'Exit Sub
' End If
' if MyReg.DeleteValue("ValName") then
'Msgbox "Value Deleted"
' else
'msgbox "Couldn't Delete Value"
' end if
' MyReg.CloseRegistry


Public Function DeleteValue(ByVal sValueName As String) As Boolean

    Dim lReturn As Long
    
    If Not OpenRegOk Then Exit Function
    
    lReturn = RegDeleteValue(hKey, sValueName)


    If lReturn = 0 Then
        DeleteValue = True
    Else
        DeleteValue = False
    End If

End Function

'This function will return a specific value from the registry
'eg.
' Dim MyString As String, MyReg As New CReadWriteEasyReg, i As In
'     teger
' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "HardWare\Descrip
'     tion\System\CentralProcessor\0") Then
'MsgBox "Couldn't open the registry"
'Exit Sub
' End If
' MyString = MyReg.GetValue("Identifier")
' Debug.Print MyString
' MyReg.CloseRegistry


Function GetValue(ByVal VarName As String, Optional ReturnBinStr As Boolean = False) As Variant

    On Error GoTo handelgetavalue
    Dim i As Integer
    Dim SubKey_Value As String, TempStr As String, ReturnArray() As Variant
    Dim length As Long
    'Dim value_type As Long
    Dim RtnVal As Long, value_Type As lDataType
    
    If Not OpenRegOk Then Exit Function
    
    'Read the size of the value value
    RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal 0&, length)


    Select Case RtnVal
        Case 0 'Ok so continue
        Case 2 'Not Found
        Exit Function
        Case 5 'Access Denied
        GetValue = "Access Denied"
        Exit Function
        Case Else 'What?
        GetValue = "RegQueryValueEx Returned : (" & RtnVal & ")"
        Exit Function
    End Select

'declare the size of the value and read it


SubKey_Value = Space$(length)

    RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal SubKey_Value, length)


    Select Case value_Type
        Case REG_NONE
        'Not defined


SubKey_Value = "Not defined value_type=REG_NONE"

    Case REG_SZ 'A null-terminated String


SubKey_Value = Left$(SubKey_Value, length - 1)

    Case REG_EXPAND_SZ
    'A null-terminated string that contains unexpanded references to
    'environment variables (for example, "%PATH%").
    'Use ExpandEnvironmentStrings to expand


SubKey_Value = Left$(SubKey_Value, length - 1)

    Case REG_BINARY 'Binary data in any form.


SubKey_Value = Left$(SubKey_Value, length)



    If Not ReturnBinStr Then
        TempStr = ""


        For i = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
        Next i



SubKey_Value = TempStr

End If

Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN 'A 32-bit number.


SubKey_Value = Left$(SubKey_Value, length)



    If Not ReturnBinStr Then
        TempStr = ""


        For i = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
        Next i



SubKey_Value = TempStr

End If

Case REG_DWORD_BIG_ENDIAN
'A 32-bit number in big-endian format.
'In big-endian format, a multi-byte value is stored in memory fro
'     m
'the highest byte (the "big end") to the lowest byte. For example
'     ,
'the value 0x12345678 is stored as (0x12 0x34 0x56 0x78) in big-e
'     ndian
'format.
Case REG_LINK
'A Unicode symbolic link. Used internally; applications should no
'     t
'use this type.


SubKey_Value = "Not defined value_type=REG_LINK"

    Case REG_MULTI_SZ
    'Array of null-terminated string


SubKey_Value = Left$(SubKey_Value, length)

    i = 0


    While Len(SubKey_Value) > 0
        ReDim Preserve ReturnArray(i) As Variant
        ReturnArray(i) = Mid$(SubKey_Value, 1, InStr(1, SubKey_Value, Chr(0)) - 1)


SubKey_Value = Mid$(SubKey_Value, InStr(1, SubKey_Value, Chr(0)) + 1)

    i = i + 1
Wend

GetValue = ReturnArray
Exit Function
Case REG_RESOURCE_LIST
'Device driver resource list.


SubKey_Value = "Not defined value_type=REG_RESOURCE_LIST"

    Case REG_FULL_RESOURCE_DESCRIPTOR
    'Device driver resource list.


SubKey_Value = "Not defined value_type=REG_FULL_RESOURCE_DESCRIPTOR"

    Case REG_RESOURCE_REQUIREMENTS_LIST
    'Device driver resource list.


SubKey_Value = "Not defined value_type=REG_RESOURCE_REQUIREMENTS_LIST"

    Case Else


SubKey_Value = "value_type=" & value_Type

End Select

GetValue = SubKey_Value
Exit Function
handelgetavalue:
GetValue = ""
Exit Function
End Function

'This property returns the current KeyValue


Public Property Get RegistryRootKey() As HKeys

    RegistryRootKey = RootHKey
End Property

'This property returns the current 'Registry Directory' your in


Public Property Get SubDirectory() As String



SubDirectory = SubDir

End Property

'This function open's the registry at a specific 'Registry Direct
'     ory'
'eg.
' Dim MyVariant As Variant, MyReg As New CReadWriteEasyReg, i As
'     Integer
' If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "") Then
'MsgBox "Couldn't open the registry"
'Exit Sub
' End If
' MyVariant = MyReg.GetAllSubDirectories
' For i = LBound(MyVariant) To UBound(MyVariant)
'Debug.Print MyVariant(i)
' Next i
' MyReg.CloseRegistry


Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Integer

    On Error GoTo OpenReg
    Dim ReturnVal As Integer


    If RtHKey = 0 Then
        OpenRegistry = False
        OpenRegOk = False
        Exit Function
    End If

    RootHKey = RtHKey


SubDir = SbDr



    If OpenRegOk Then
        CloseRegistry
        OpenRegOk = False
    End If

    ReturnVal = RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_READ, hKey)


    If ReturnVal <> 0 Then
        OpenRegistry = False
        Exit Function
    End If

    OpenRegOk = True
    OpenRegistry = True
    Exit Function
OpenReg:
    OpenRegOk = False
    OpenRegistry = False
    Exit Function
End Function



Public Function OneBackOnKey()



SubDir = Mid$(SubDir, 1, FindLastBackSlash(SubDir) - 1)

    CloseRegistry
    OpenRegistry RootHKey, SubDir
End Function

'This function should be called after you're done with the regist
'     ry
'eg. (see other examples)


Public Function CloseRegistry() As Boolean

    On Error Resume Next


    If RegCloseKey(hKey) <> 0 Then
        CloseRegistry = False
        Exit Function
    End If

    CloseRegistry = True
    OpenRegOk = False
End Function



Private Sub Class_Initialize()

    RootHKey = &H0


SubDir = ""

    hKey = 0
    OpenRegOk = False
End Sub



Private Sub Class_Terminate()

    On Error Resume Next


    If RegCloseKey(hKey) <> 0 Then
        Exit Sub
    End If

End Sub



Public Function SortArrayAscending(ValueList As Variant) As Variant

    On Error GoTo handelsort
    Dim RipVal As Variant
    Dim RipOrdinal As Long
    Dim RipDescent As Long
    Dim PrivateBuffer As Variant
    Dim Placed As Boolean
    Dim x As Long
    Dim y As Long


    If IsArray(ValueList) Then
        PrivateBuffer = ValueList
        'Ok, we start at the second position in the array and go
        'from there
        RipOrdinal = 1
        RipDescent = 1


        For y = 1 To UBound(PrivateBuffer)
            RipVal = PrivateBuffer(y)
            If y <> 1 Then RipDescent = y


            Do Until Placed


                If PrivateBuffer(RipDescent - 1) >= RipVal Then
                    RipDescent = RipDescent - 1


                    If RipDescent = 0 Then


                        For x = y To RipDescent Step -1
                            If x = 0 Then Exit For
                            PrivateBuffer(x) = PrivateBuffer(x - 1)
                        Next x

                        PrivateBuffer(RipDescent) = RipVal
                        Placed = True
                    End If

                Else
                    'shift the array to the right


                    For x = y To RipDescent Step -1
                        If x = 0 Then Exit For
                        PrivateBuffer(x) = PrivateBuffer(x - 1)
                    Next x

                    'insert the ripped value
                    PrivateBuffer(RipDescent) = RipVal
                    Placed = True
                End If

            Loop

            Placed = False
        Next y

        SortArrayAscending = PrivateBuffer
    Else
        SortArrayAscending = ValueList
    End If

    Exit Function
handelsort:
    SortArrayAscending = ValueList
    Exit Function
End Function



Private Function FindLastBackSlash(VarValue As Variant) As Integer

    Dim i As Integer, iRtn As Integer
    iRtn = 0


    For i = Len(VarValue) To 1 Step -1


        If Mid$(VarValue, i, 1) = "\" Then
            iRtn = i
            Exit For
        End If

    Next i

    FindLastBackSlash = iRtn
End Function



