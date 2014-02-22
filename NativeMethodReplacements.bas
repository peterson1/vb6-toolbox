Attribute VB_Name = "Replacements"
Option Explicit
Const MODULE_NAME$ = "Replacements"

Private Enum ReplacementsErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As ReplacementsErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal ptr As Long, Value As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, Optional lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Private m_A() As Long
Private m_AP As Long
Private m_H(0 To 6) As Long
Private m_HP As Long


'Public Function ArrayEmpty(anyArray1D() As String) As Boolean
'    On Error GoTo ErrH
'
'    'If (Not anyArray1D) = -1 Then ArrayEmpty = True
'    If SafeArrayGetDim(anyArray1D) = 0 Then ArrayEmpty = True
'
'ErrH: Blame "ArrayEmpty"
'End Function


Public Function GetTokens(strWithTokens As String _
                        , startMarkr As String _
                        , endMarkr As String _
                        , Optional compareMethd As VbCompareMethod = vbBinaryCompare _
                        ) As String()
    Dim occurens&, ss$(), i&, startPos&, endPos&
    On Error GoTo Cleanup
    
    ' validations
    ErrorIf LenB(startMarkr) = 0, "Token start marker should not be _blank_."
    ErrorIf LenB(endMarkr) = 0, "Token end marker should not be _blank_."
    
    occurens = Occurence(strWithTokens, startMarkr, compareMethd)
    ErrorIf occurens <> Occurence(strWithTokens, endMarkr, compareMethd) _
         , "Count of start & end markers do not match."
    
    ReDim ss(occurens - 1)
    
    For i = 0 To UBound(ss)
        If startPos = 0 Then
            startPos = InStr(1, strWithTokens, startMarkr, compareMethd)
        Else
            startPos = InStr(endPos + Len(endMarkr), strWithTokens, startMarkr, compareMethd)
        End If
        
        startPos = startPos + Len(startMarkr)
        
        endPos = InStr(startPos, strWithTokens, endMarkr, compareMethd)
        
        ss(i) = Mid$(strWithTokens, startPos, endPos - startPos)
    Next i
    
    GetTokens = ss
    
Cleanup:    SaveError
            Erase ss
  LoadError "GetTokens" ', "details of error"
End Function


Public Property Get InText(strText As String _
                         , findThisStr As String _
                         , Optional compareMethd As VbCompareMethod = vbBinaryCompare _
                         ) As Boolean
    On Error GoTo ErrH
    
    InText = InStrRev(strText, findThisStr, , compareMethd) <> 0
    
ErrH: Blame "InText"
End Property


'  Counts the number of times a string appears in a text.
'
Public Property Get Occurence(strText As String _
                            , findThisStr As String _
                            , Optional compareMethd As VbCompareMethod = vbBinaryCompare _
                            ) As Long
    Dim ss$()
    On Error GoTo Cleanup
    
    If InStrRev(strText, findThisStr, , compareMethd) = 0 Then Exit Property
    
    ss = Split(strText, findThisStr, , compareMethd)
    
    Occurence = UBound(ss)
    
Cleanup:    SaveError
            Erase ss
  LoadError "Occurence" ', "details of error"
End Property

Public Property Get Trim(untrimmedTxt As String) As String
    Dim startPos&, endPos&
    On Error GoTo ErrH
    
    If LenB(untrimmedTxt) = 0 Then Exit Property
    
    Do
        startPos = startPos + 1
    Loop While Trimmable(Mid$(untrimmedTxt, startPos, 1))
    
    
    endPos = Len(untrimmedTxt)
    Do While Trimmable(Mid$(untrimmedTxt, endPos, 1))
        endPos = endPos - 1
        If endPos = 0 Then
            endPos = startPos
            Exit Do
        End If
    Loop
    
    
    Trim = Mid$(untrimmedTxt, startPos, endPos - startPos + 1)
    
ErrH: Blame "Trim"
End Property

Public Property Get Trimmable(singleCharactr As String) As Boolean
    If LenB(singleCharactr) = 0 Then Exit Property
    Select Case Asc(singleCharactr)
    
        Case 32:    Trimmable = True ' space
        Case 13:    Trimmable = True ' vbCr / vbNewLine
        Case 9:     Trimmable = True ' vbTab
        Case 10:    Trimmable = True ' vbLf
        
    End Select
End Property


Public Property Get ApplicationVersion() As String
    With App
        ApplicationVersion = .Title & "  v" & .Major _
                                      & "." & .Minor _
                                      & "." & .Revision
    End With
End Property


Public Function Shell(ByVal pathOfApplication As String _
                    , Optional windwStyle As VbAppWinStyle = vbNormalFocus _
                    ) As Double
    
    ' quote if unquoted
    If Left$(pathOfApplication, 1) <> """" Then _
        pathOfApplication = """" & pathOfApplication & """"
    
    Shell = VBA.Shell(pathOfApplication, windwStyle)
End Function



Public Function AllocString(ByVal ptr As Long, ByVal lngth As Long) As String
    On Error GoTo Cleanup
    
    Procedure(AddressOf Replacements.AllocString) = API("oleaut32", "SysAllocStringByteLen")
    AllocString = Replacements.AllocString(ptr, lngth)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "AllocString" ', "details of error"
End Function

Public Sub PutLong(ByVal ptr As Long, ByVal val As Long)
    Procedure(AddressOf Replacements.PutLong) = API("msvbvm60", "PutMem4")
    PutLong ptr, val
End Sub

Public Sub RtlMove(ByVal dest As Long, ByVal src As Long, ByVal lngth As Long)
    Procedure(AddressOf Replacements.RtlMove) = API("kernel32", "RtlMoveMemory")
    RtlMove dest, src, lngth
End Sub

Private Function InIDE(Optional IDE) As Boolean
    If IsMissing(IDE) Then Debug.Assert Not InIDE(InIDE) Else IDE = True
End Function

Private Property Get Procedure(ByVal addressOfDest As Long) As Long
    ' get correct pointer to procedure in IDE
    If Not InIDE Then Procedure = addressOfDest Else GetMem4 addressOfDest + &H16&, Procedure
End Property

Private Property Let Procedure(ByVal addressOfDest As Long, ByVal addressOfSrc As Long)
    Dim jmp As Currency, pid&
    ' get process handle
    pid = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)
    If pid Then
        ' get correct pointer to procedure in IDE
        If InIDE Then GetMem4 addressOfDest + &H16&, addressOfDest
        Debug.Assert App.hInstance
        ' ASM JMP (0xE9) followed by bytes to jump in memory
        jmp = (&HE9& * 0.0001@) + (addressOfSrc - addressOfDest - 5@) * 0.0256@
        ' write the JMP over the destination procedure
        WriteProcessMemory pid, ByVal addressOfDest, jmp, 5
        ' close process handle
        CloseHandle pid
    End If
End Property

Private Property Get API(moduleNme As String, procName As String) As Long
    Dim handl&
    On Error GoTo Cleanup
    
    handl = GetModuleHandle(moduleNme)
    If handl = 0 Then handl = LoadLibrary(moduleNme)
    If handl Then API = GetProcAddress(handl, procName)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "API" ', "details of error"
End Property




'  modified from:
'   http://www.vbforums.com/showthread.php?540323-VB6-Faster-Split-amp-Join-(development)
'
Public Function Join(strArray() As String _
                   , Optional delimtr$ = vbNullString _
                   ) As String
    Static hi1&(0 To 6), hpi1&
    Static hi2&(0 To 6), hpi2&
    Static hl1&(0 To 6), hpl1&
    Static hl2&(0 To 6), hpl2&
    Static hl3&(0 To 6), hpl3&
    Static hl4&(0 To 6), hpl4&
    Static hs&(0 To 6), hps&
    Static hsi&(0 To 6), hpsi&
    Static hsl&(0 To 6), hpsl&

    Dim i1%(), api1&, di1%
    Dim i2%(), api2&
    Dim l1&(), apl1&, dl1&
    Dim l2&(), apl2&, dl2&
    Dim l3&(), apl3&, dl3&
    Dim l4&(), apl4&, dl4&
    Dim s&(), aps&, ps&
    Dim si%(), apsi&, pi&
    Dim sl&(), apsl&, pl&
    
    Dim j&, l&, ld&, pc&, pd&, t&, us&, p&, ptr&, retrnTo&
    On Error GoTo Cleanup
    
    ' has the array been initialized in the first place?
    p = Not Not strArray
    Debug.Assert App.hInstance
    If p Then
        ' must be zero base
        If LBound(strArray) Then Exit Function
        ' upper bound check...
        us = UBound(strArray)
        Select Case us
        ' must have at least one element!
        Case -1: Exit Function
        ' only one element?
        Case 0: Join = strArray(0): Exit Function
        End Select
        ' calculate total length
        For j = 0 To us
            t = t + LenB(strArray(j))
        Next j
        ' get delimiter length
        ld = LenB(delimtr)
        ' because you can be crazy on optimizing performance... including rare special cases
        If t Then
            ' add delimiters' length to total length
            If ld Then t = t + ld * us
        ElseIf ld = 2 Then
            ' single character replicate
            Join = String$(us, delimtr)
        ElseIf ld Then
            ' half character or multicharacter replicate
            Join = AllocString(0, ld * us)
            p = 1
            ' join the first four bytes and until we have even length
            Do
                MidB$(Join, p, ld) = delimtr
                p = p + ld
            ' make also sure we do not run out of the string space
            Loop While (p <= LenB(Join)) And ((p < 4) Or (p And 1) = 0)
            ' now use a faster replicate for the rest
            If p <= LenB(Join) Then MidB$(Join, p) = Join
        End If
        ' so do we join?
        If t Then
            ' generic safe array hack
            If m_AP = 0 Then
                ' array variable pointer
                m_AP = ArrPtr(m_A)
                ' create a safe array header
                m_H(0) = vbLong: m_H(1) = &H800001: m_H(2) = 4: m_H(5) = &H7FFFFFFF
                ' header pointer
                m_HP = VarPtr(m_H(1))
            End If
            ' set pointer to array
            PutLong m_AP, m_HP
            ' local safe array hack
            api1 = ArrPtr(i1)
            api2 = ArrPtr(i2)
            apl1 = ArrPtr(l1)
            apl2 = ArrPtr(l2)
            apl3 = ArrPtr(l3)
            apl4 = ArrPtr(l4)
            aps = ArrPtr(s)
            apsi = ArrPtr(si)
            apsl = ArrPtr(sl)
            If hps = 0 Then
                ' Join string
                hi1(0) = vbInteger: hi1(1) = &H800001: hi1(2) = 2: hi1(5) = &H7FFFFFFF
                hi2(0) = vbInteger: hi2(1) = &H800001: hi2(2) = 2: hi2(5) = &H7FFFFFFF
                hl1(0) = vbLong: hl1(1) = &H800001: hl1(2) = 4: hl1(5) = &H3FFFFFFF
                hl2(0) = vbLong: hl2(1) = &H800001: hl2(2) = 4: hl2(5) = &H3FFFFFFF
                hl3(0) = vbLong: hl3(1) = &H800001: hl3(2) = 4: hl3(5) = &H3FFFFFFF
                hl4(0) = vbLong: hl4(1) = &H800001: hl4(2) = 4: hl4(5) = &H3FFFFFFF
                ' string pointers
                hs(0) = vbLong: hs(1) = &H800001: hs(2) = 4
                ' strArray's current string
                hsi(0) = vbInteger: hsi(1) = &H800001: hsi(2) = 2: hsi(5) = &H7FFFFFFF
                hsl(0) = vbLong: hsl(1) = &H800001: hsl(2) = 4: hsl(5) = &H3FFFFFFF
                hpi1 = VarPtr(hi1(1))
                hpi2 = VarPtr(hi2(1))
                hpl1 = VarPtr(hl1(1))
                hpl2 = VarPtr(hl2(1))
                hpl3 = VarPtr(hl3(1))
                hpl4 = VarPtr(hl4(1))
                hps = VarPtr(hs(1))
                hpsi = VarPtr(hsi(1))
                hpsl = VarPtr(hsl(1))
            End If
            ' allocate string
            Join = AllocString(0, t)
            ptr = StrPtr(Join)
            ' update S to point to strArray
            hs(4) = VarPtr(strArray(0)): hs(5) = us + 1
            ' update array pointers
            hi1(4) = ptr
            hi2(4) = ptr + 1
            hl1(4) = ptr
            hl2(4) = ptr + 1
            hl3(4) = ptr + 2
            hl4(4) = ptr + 3
            ' start safe array hack for local arrays
            m_H(4) = api1: m_A(0) = hpi1
            m_H(4) = api2: m_A(0) = hpi2
            m_H(4) = apl1: m_A(0) = hpl1
            m_H(4) = apl2: m_A(0) = hpl2
            m_H(4) = apl3: m_A(0) = hpl3
            m_H(4) = apl4: m_A(0) = hpl4
            m_H(4) = aps: m_A(0) = hps
            m_H(4) = apsi: m_A(0) = hpsi
            m_H(4) = apsl: m_A(0) = hpsl
            ' zero position
            p = 0
            pd = StrPtr(delimtr)
            ' delimiter length?
            Select Case ld
            Case Is > 16
                GoTo Over8
            ' optimized zero length join
            Case 0
                retrnTo = 0
                For j = 0 To us
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
0:                  End If
                Next j
            Case 1, 2
                hsi(4) = pd
                di1 = si(0)
                retrnTo = 1
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
1:                  End If
                    ' copy delimiter
                    pi = p \ 2
                    Select Case p And 1
                        Case 0: i1(pi) = di1
                        Case 1: i2(pi) = di1
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 3, 4
                m_H(4) = pd
                dl1 = m_A(0)
                retrnTo = 2
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
2:                  End If
                    ' copy delimiter
                    pl = p \ 4
                    Select Case p And 3
                        Case 0: l1(pl) = dl1
                        Case 2: l3(pl) = dl1
                        Case 1: l2(pl) = dl1
                        Case 3: l4(pl) = dl1
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 5, 6
                hsi(4) = pd
                di1 = si(0)
                m_H(4) = pd + 2
                dl1 = m_A(0)
                retrnTo = 3
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
3:                  End If
                    ' copy delimiter
                    pi = p \ 2
                    Select Case p And 1
                        Case 0: i1(pi) = di1
                        Case 1: i2(pi) = di1
                    End Select
                    pl = (p + 2) \ 4
                    Select Case p And 3
                        Case 2: l1(pl) = dl1
                        Case 0: l3(pl) = dl1
                        Case 3: l2(pl) = dl1
                        Case 1: l4(pl) = dl1
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 7, 8
                m_H(4) = pd
                dl1 = m_A(0)
                dl2 = m_A(1)
                retrnTo = 4
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
4:                  End If
                    ' copy delimiter
                    pl = p \ 4
                    Select Case p And 3
                        Case 0: l1(pl) = dl1: l1(pl + 1) = dl2
                        Case 2: l3(pl) = dl1: l3(pl + 1) = dl2
                        Case 1: l2(pl) = dl1: l2(pl + 1) = dl2
                        Case 3: l4(pl) = dl1: l4(pl + 1) = dl2
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 9, 10
                hsi(4) = pd
                di1 = si(0)
                m_H(4) = pd + 2
                dl1 = m_A(0)
                dl2 = m_A(1)
                retrnTo = 5
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
5:                  End If
                    ' copy delimiter
                    pi = p \ 2
                    Select Case p And 1
                        Case 0: i1(pi) = di1
                        Case 1: i2(pi) = di1
                    End Select
                    pl = (p + 2) \ 4
                    Select Case p And 3
                        Case 2: l1(pl) = dl1: l1(pl + 1) = dl2
                        Case 0: l3(pl) = dl1: l3(pl + 1) = dl2
                        Case 3: l2(pl) = dl1: l2(pl + 1) = dl2
                        Case 1: l4(pl) = dl1: l4(pl + 1) = dl2
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 11, 12
                m_H(4) = pd
                dl1 = m_A(0)
                dl2 = m_A(1)
                dl3 = m_A(2)
                retrnTo = 6
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
6:                  End If
                    ' copy delimiter
                    pl = p \ 4
                    Select Case p And 3
                        Case 0: l1(pl) = dl1: l1(pl + 1) = dl2: l1(pl + 2) = dl3
                        Case 2: l3(pl) = dl1: l3(pl + 1) = dl2: l3(pl + 2) = dl3
                        Case 1: l2(pl) = dl1: l2(pl + 1) = dl2: l2(pl + 2) = dl3
                        Case 3: l4(pl) = dl1: l4(pl + 1) = dl2: l4(pl + 2) = dl3
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 13, 14
                hsi(4) = pd
                di1 = si(0)
                m_H(4) = pd + 2
                dl1 = m_A(0)
                dl2 = m_A(1)
                dl3 = m_A(2)
                retrnTo = 7
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
7:                  End If
                    ' copy delimiter
                    pi = p \ 2
                    Select Case p And 1
                        Case 0: i1(pi) = di1
                        Case 1: i2(pi) = di1
                    End Select
                    pl = (p + 2) \ 4
                    Select Case p And 3
                        Case 2: l1(pl) = dl1: l1(pl + 1) = dl2: l1(pl + 2) = dl3
                        Case 0: l3(pl) = dl1: l3(pl + 1) = dl2: l3(pl + 2) = dl3
                        Case 3: l2(pl) = dl1: l2(pl + 1) = dl2: l2(pl + 2) = dl3
                        Case 1: l4(pl) = dl1: l4(pl + 1) = dl2: l4(pl + 2) = dl3
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case 15, 16
                m_H(4) = pd
                dl1 = m_A(0)
                dl2 = m_A(1)
                dl3 = m_A(2)
                dl4 = m_A(3)
                retrnTo = 8
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
8:                  End If
                    ' copy delimiter
                    pl = p \ 4
                    Select Case p And 3
                        Case 0: l1(pl) = dl1: l1(pl + 1) = dl2: l1(pl + 2) = dl3: l1(pl + 3) = dl4
                        Case 2: l3(pl) = dl1: l3(pl + 1) = dl2: l3(pl + 2) = dl3: l3(pl + 3) = dl4
                        Case 1: l2(pl) = dl1: l2(pl + 1) = dl2: l2(pl + 2) = dl3: l2(pl + 3) = dl4
                        Case 3: l4(pl) = dl1: l4(pl + 1) = dl2: l4(pl + 2) = dl3: l4(pl + 3) = dl4
                    End Select
                    ' update position
                    p = p + ld
                Next j
            Case Else
Over8:          retrnTo = 9
                For j = 0 To us - 1
                    ' current length
                    l = LenB(strArray(j))
                    If l Then
                        ' current pointer
                        pc = s(j)
                        ' copy from StringArray
                        GoTo CopyRoutine
9:                  End If
                    ' copy delimiter
                    RtlMove ptr + p, pd, ld
                    ' update position
                    p = p + ld
                Next j
            End Select
            ' last copy?
            If j = us Then
                ' current length
                l = LenB(strArray(j))
                If l Then
                    ' current pointer
                    pc = s(j)
                    retrnTo = 10
                    ' copy from StringArray (last item)
                    GoTo CopyRoutine
10:              End If
            End If
            ' end safe array hacks
            m_H(4) = api1: m_A(0) = 0
            m_H(4) = api2: m_A(0) = 0
            m_H(4) = apl1: m_A(0) = 0
            m_H(4) = apl2: m_A(0) = 0
            m_H(4) = apl3: m_A(0) = 0
            m_H(4) = apl4: m_A(0) = 0
            m_H(4) = aps: m_A(0) = 0
            m_H(4) = apsi: m_A(0) = 0
            m_H(4) = apsl: m_A(0) = 0
            m_H(4) = m_AP: m_A(0) = 0
        End If
    End If
    Exit Function
CopyRoutine:
    Select Case l
    Case Is > 16
        RtlMove ptr + p, pc, l
        GoTo ReturnPoint
    Case 1, 2
        hsi(4) = pc
        pi = p \ 2
        Select Case p And 1
        Case 0: i1(pi) = si(0)
        Case 1: i2(pi) = si(0)
        End Select
        GoTo ReturnPoint
    Case 3, 4
        hsl(4) = pc
        pl = p \ 4
        Select Case p And 3
        Case 0: l1(pl) = sl(0)
        Case 2: l3(pl) = sl(0)
        Case 1: l2(pl) = sl(0)
        Case 3: l4(pl) = sl(0)
        End Select
        GoTo ReturnPoint
    Case 5, 6
        hsi(4) = pc
        pi = p \ 2
        Select Case p And 1
        Case 0: i1(pi) = si(0): i1(pi + 1) = si(1): i1(pi + 2) = si(2)
        Case 1: i2(pi) = si(0): i2(pi + 1) = si(1): i2(pi + 2) = si(2)
        End Select
        GoTo ReturnPoint
    Case 7, 8
        hsl(4) = pc
        pl = p \ 4
        Select Case p And 3
        Case 0: l1(pl) = sl(0): l1(pl + 1) = sl(1)
        Case 2: l3(pl) = sl(0): l3(pl + 1) = sl(1)
        Case 1: l2(pl) = sl(0): l2(pl + 1) = sl(1)
        Case 3: l4(pl) = sl(0): l4(pl + 1) = sl(1)
        End Select
        GoTo ReturnPoint
    Case 9, 10
        hsi(4) = pc
        pi = p \ 2
        Select Case p And 1
        Case 0: i1(pi) = si(0)
        Case 1: i2(pi) = si(0)
        End Select
        hsl(4) = pc + 2
        pl = (p + 2) \ 4
        Select Case (p + 2) And 3
        Case 0: l1(pl) = sl(0): l1(pl + 1) = sl(1)
        Case 2: l3(pl) = sl(0): l3(pl + 1) = sl(1)
        Case 1: l2(pl) = sl(0): l2(pl + 1) = sl(1)
        Case 3: l4(pl) = sl(0): l4(pl + 1) = sl(1)
        End Select
        GoTo ReturnPoint
    Case 11, 12
        hsl(4) = pc
        pl = p \ 4
        Select Case p And 3
        Case 0: l1(pl) = sl(0): l1(pl + 1) = sl(1): l1(pl + 2) = sl(2)
        Case 2: l3(pl) = sl(0): l3(pl + 1) = sl(1): l3(pl + 2) = sl(2)
        Case 1: l2(pl) = sl(0): l2(pl + 1) = sl(1): l2(pl + 2) = sl(2)
        Case 3: l4(pl) = sl(0): l4(pl + 1) = sl(1): l4(pl + 2) = sl(2)
        End Select
        GoTo ReturnPoint
    Case 13, 14
        hsi(4) = pc
        pi = p \ 2
        Select Case p And 1
        Case 0: i1(pi) = si(0)
        Case 1: i2(pi) = si(0)
        End Select
        hsl(4) = pc + 2
        pl = (p + 2) \ 4
        Select Case (p + 2) And 3
        Case 0: l1(pl) = sl(0): l1(pl + 1) = sl(1): l1(pl + 2) = sl(2)
        Case 2: l3(pl) = sl(0): l3(pl + 1) = sl(1): l3(pl + 2) = sl(2)
        Case 1: l2(pl) = sl(0): l2(pl + 1) = sl(1): l2(pl + 2) = sl(2)
        Case 3: l4(pl) = sl(0): l4(pl + 1) = sl(1): l4(pl + 2) = sl(2)
        End Select
        GoTo ReturnPoint
    Case 15, 16
        hsl(4) = pc
        pl = p \ 4
        Select Case p And 3
        Case 0: l1(pl) = sl(0): l1(pl + 1) = sl(1): l1(pl + 2) = sl(2): l1(pl + 3) = sl(3)
        Case 2: l3(pl) = sl(0): l3(pl + 1) = sl(1): l3(pl + 2) = sl(2): l3(pl + 3) = sl(3)
        Case 1: l2(pl) = sl(0): l2(pl + 1) = sl(1): l2(pl + 2) = sl(2): l2(pl + 3) = sl(3)
        Case 3: l4(pl) = sl(0): l4(pl + 1) = sl(1): l4(pl + 2) = sl(2): l4(pl + 3) = sl(3)
        End Select
        GoTo ReturnPoint
    End Select
ReturnPoint:
    ' update position
    p = p + l
    ' where we go back?
    Select Case retrnTo
    Case 0: GoTo 0:     Case 1: GoTo 1:     Case 2: GoTo 2:     Case 3: GoTo 3
    Case 4: GoTo 4:     Case 5: GoTo 5:     Case 6: GoTo 6:     Case 7: GoTo 7
    Case 8: GoTo 8:     Case 9: GoTo 9:     Case 10: GoTo 10
    End Select
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Join" ', "details of error"
End Function



Public Function IsNumeric(exprssion As Variant _
                        ) As Boolean
    
    
    '  if native rejects it, so should we
    '
    If Not VBA.IsNumeric(exprssion) Then Exit Function
    
    
    '  native sees Empty as numeric
    '   - we should disagree
    '
    If IsEmpty(exprssion) Then Exit Function
    
    
    '  more native bugs
    '
    If Left(exprssion, 1) = "+" Then Exit Function
    If Right(exprssion, 1) = "+" Then Exit Function
    If Right$(exprssion, 1) = "-" Then Exit Function
    
    
    '  if starts with "0",
    '   - treat as non-numeric to preserve zero-pad
    '   - unless it's a decimal (contains ".")
    '
    If Len(exprssion) <> 1 Then
        If Left(exprssion, 1) = "0" Then
            If InStr(exprssion, ".") = 0 Then Exit Function
        End If
    End If
    
    If InStr(1, exprssion, "e", vbTextCompare) Then
        If InStr(2, exprssion, "+") = 0 Then Exit Function
    End If
    
    If InStr(1, exprssion, "d", vbTextCompare) Then Exit Function
    
    If InStr(exprssion, ",") Then Exit Function
    
    IsNumeric = True
End Function



Public Function SplitTrim(strText As String _
                        , Optional delimitr As Variant _
                        , Optional limit As Long = -1 _
                        , Optional compareMethd As VbCompareMethod = vbBinaryCompare _
                        ) As String()
    Dim i&, ss$()
    On Error GoTo Cleanup
    
    ss = Split(strText, delimitr, limit, compareMethd)
    
    For i = 0 To UBound(ss)
        ss(i) = Trim$(ss(i))
    Next i
    
    SplitTrim = ss
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "SplitTrim" ', "details of error"
End Function









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As ReplacementsErrors = -1 _
                  )
    If errCondition Then Err.Raise errorNumbr, MODULE_NAME, errorMsg
End Sub

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As ReplacementsErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub SaveError()
    With mError
        If Err Then
            .HasError = True
            .Description = Err.Description
            .Number = Err.Number
            .Source = Err.Source
            
        Else
            .HasError = False
            .Description = vbNullString
            .Number = 0
            .Source = vbNullString
        End If
    End With
    Err.Clear
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As ReplacementsErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = MODULE_NAME & "." & currntProcedure & "()"

        If errorNumbr = -1 Then errorNumbr = .Number
        
        Select Case errorNumbr
            Case NotInitted
                errorDescrption = MODULE_NAME & " not initted." & vbCrLf _
                               & "Please call " & MODULE_NAME _
                               & ".Init() before " & currntProcedure & "."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub

