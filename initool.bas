Attribute VB_Name = "iniTool"
Sub PutValue(putcat, putvar, putval, putfile)
Dim fileCol(1 To 9000) As String
Dim foundCat As Boolean
Dim foundVar As Boolean
Dim catPos As Integer
Dim varPos As Integer
fnum = FreeFile
putcat = Trim(putcat)
putcat = LCase(putcat)
putfile = Trim(putfile)
putfile = LCase(putfile)
putvar = LCase(putvar)
putvar = Trim(putvar)
putval = LCase(putval)
putval = Trim(putval)
If Dir(putfile) = "" Then
    Open putfile For Append As #fnum
    Close #fnum
End If

Open putfile For Input As #fnum
    Do While Not EOF(fnum)
        DoEvents
        Counter = Counter + 1
        Line Input #fnum, l1
        fileCol(Counter) = l1
    Loop
Close #fnum

For i = 1 To Counter
    DoEvents
    If InStr(LCase(fileCol(i)), "[" & putcat & "]") <> 0 Then
        foundCat = True
        catPos = i
        For x = i To Counter
            DoEvents
            If InStr(fileCol(x), "[") <> 0 And LCase(fileCol(x)) <> "[" & putcat & "]" Then Exit For
            If InStr(LCase(fileCol(x)), putvar & "=") <> 0 Then
                foundVar = True
                varPos = x
            End If
        Next x
    End If
Next i

If foundCat = True And foundVar = True Then
fileCol(varPos) = putvar & "=" & putval
Kill putfile
Open putfile For Append As #fnum
    For i = 1 To Counter
        Print #fnum, fileCol(i)
        DoEvents
    Next i
Close #fnum
Exit Sub
End If


If foundCat = True And foundVar = False Then
Kill putfile
Open putfile For Append As #fnum
    For i = 1 To Counter
        Print #fnum, fileCol(i)
        If i = catPos Then Print #fnum, putvar & "=" & putval
    Next i
Close #fnum
Exit Sub
End If

If foundCat = False And foundVar = False Then
Kill putfile
Open putfile For Append As #fnum
    For i = 1 To Counter
        Print #fnum, fileCol(i)
    Next i
    Print #fnum, "[" & putcat & "]"
    Print #fnum, putvar & "=" & putval
Close #fnum
End If

End Sub
Function GetValue(getcat, getfield, getfile) As String
'example usage:
'username = GetValue("UserInfo", "Username", "myprog.ini")

If Dir(getfile) = "" Then Exit Function
getcat = LCase(getcat)
getfield = LCase(getfield)
fnum = FreeFile
Open getfile For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, l1
        l1 = Trim(l1)
        l1 = LCase(l1)
        If InStr(l1, "[") <> 0 Then
            If LCase(Mid(l1, (InStr(l1, "[") + 1), (Len(l1) - 2))) = getcat Then
                Do Until EOF(fnum) Or l2 = "["
                    Line Input #fnum, l2
                    l2 = Trim(l2)
                    If InStr(l2, "]") <> 0 Then
                        Close fnum
                        Exit Function
                    End If
                    If InStr(l2, "=") <> 0 Then
                        If LCase(Left(l2, (InStr(l2, "=") - 1))) = getfield Then
                            GetValue = Trim(Mid(l2, InStr(l2, "=") + 1, Len(l2)))
                            Close fnum
                            Exit Function
                        End If
                    End If
                Loop
            End If
        End If
    Loop
    Close fnum
End Function

