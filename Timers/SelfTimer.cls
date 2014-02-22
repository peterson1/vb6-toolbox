VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelfTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* SelfTimer 1.1 - Timer class module
'* ----------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) Free for commercial and non-commercial usage.
'*        4) Use at your own risk. No support guaranteed.
'*
'* REQUIREMENTS
'* ------------
'* Huh what? Just this one class module. No extra files required.
'*
'* HOW TO ADD TO YOUR PROGRAM
'* --------------------------
'* 1) Copy SelfTimer.cls to your project folder.
'* 2) In your project, add SelfTimer.cls
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.1 (2008-06-23)
'* - Simplified even further thanks to Paul's improved code. All Sc functions removed, everything
'*   SelfCallback related is in Private_Start and Private_Stop.
'*
'* Version 1.0 (2008-06-15)
'* - It was requested to simplify SelfCallback code and make a one class module timer at VBForums.
'*   I renamed and cleaned up the code to fit my own coding style. Good or bad, a matter of taste.
'*
'* CREDITS
'* -------
'* Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback
'*************************************************************************************************
Option Explicit

' events
Public Event Timer(ByVal Seconds As Currency)

' public properties
Private m_Enabled As Boolean
Private m_Interval As Long

' private variables
Private m_Start As Currency
Private m_TimerID As Long
Private m_TimerProc As Long

Private Declare Sub GetMem4LngToCur Lib "msvbvm60" Alias "GetMem4" (ByRef lngVar As Long, curVar As Currency)
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

' the following are the requirements for SelfCallback
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, retVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal addr As Long, retVal As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal addr As Long, ByVal newVal As Integer)
Private Declare Sub RtlMachineCodeCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As MachineCode, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

' self-documentation: less comments and constants
Private Type MachineCode        ' 37 * 4 = 148 bytes
    ' thunk
    OwnerPtr As Long            ' 0
    CallbackAddress As Long     ' 1
    API_EbMode As Long          ' 2
    API_IsBadCodePtr As Long    ' 3
    API_KillTimer As Long       ' 4
    ' code
    MC1(5 To 5) As Long         ' 5
    AllocatedDataPtr As Long    ' 6
    MC2(7 To 17) As Long        ' 7 - 17
    ParamCount As Long          ' 18
    MC3(19 To 35) As Long       ' 19 - 35
    Ordinal As Long             ' 36
End Type

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
    m_Enabled = NewValue
    ' we can call these both; the order is important, of course
    Private_Stop
    Private_Start
End Property
Public Property Get Interval() As Long
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal NewValue As Long)
    ' validate range
    If NewValue >= 0 Then
        ' kill existing timer?
        Private_Stop
        ' update value
        m_Interval = NewValue
        ' start timer
        Private_Start
    End If
End Property
Private Sub Private_Start()
    Dim bytValue As Byte, bytSignature As Byte, lngA As Long
    Dim lngAddress As Long, lngMethod As Long, lngObject As Long
    Dim lngCallback As Long, lngMCmem As Long, udtMC As MachineCode

    ' index of ordinal (always 1 in this class, the last procedure)
    Const Ordinal As Long = 1
    ' number of parameters (for TimerProc this is always 4, thus this is a constant)
    Const ParamCount As Long = 4

    ' start only if enabled, interval is set and we have not done this already
    If m_TimerProc = 0 And m_Enabled And m_Interval <> 0 Then
        ' get object pointer
        lngObject = ObjPtr(Me)
        ' get VTable address
        GetMem4 lngObject, lngAddress
        ' Class method (see SelfSub code for non-Class values)
        lngAddress = lngAddress + &H1C&
        ' get method pointer
        GetMem4 lngAddress, lngMethod
        ' get method signature byte: &H33 = pseudo-code, &HE9 = native code
        GetMem1 lngMethod, bytSignature
        ' next VTable address
        lngAddress = lngAddress + 4&
        ' try a "reasonable" amount of VTable entries
        For lngA = 511 To 1 Step -1
            ' get method pointer
            GetMem4 lngAddress, lngMethod
            ' see if we are out of VTable (I use "Then Else" because True conditions are faster)
            If IsBadCodePtr(lngMethod) = 0& Then Else Exit For
            ' get method signature byte
            GetMem1 lngMethod, bytValue
            ' if it is invalid we are out of VTable
            If bytValue = bytSignature Then Else Exit For
            ' try next one
            lngAddress = lngAddress + 4&
        Next lngA
        ' if lngA = 0 we looped through the entire loop; if that did not happen, we get the pointer
        If lngA Then GetMem4 lngAddress - (Ordinal * 4&), lngCallback
        ' verify we got the TimerProc callback address of ordinal 1
        If lngCallback Then
            ' allocate executable memory
            lngMCmem = VirtualAlloc(0, LenB(udtMC), &H1000&, &H40&) 'Length, MEM_COMMIT, PAGE_RWX
            ' verify we got it
            If lngMCmem Then
                With udtMC
                    ' thunk
                    .OwnerPtr = lngObject
                    .CallbackAddress = lngCallback
                    If App.LogMode = 0 Then
                        ' for IDE safety, store the EbMode function address in the thunk data
                        .API_EbMode = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")
                    End If
                    .API_IsBadCodePtr = GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
                    .API_KillTimer = GetProcAddress(GetModuleHandleA("user32"), "KillTimer")
                    ' actual machine code
                    .MC1(5&) = &HBB60E089:    .MC2(7&) = &H73FFC589
                    .MC2(8&) = &HC53FF04:     .MC2(9&) = &H59E80A74
                    .MC2(10) = &HE9000000:    .MC2(11) = &H30&
                    .MC2(12) = &H87B81:       .MC2(13) = &H75000000
                    .MC2(14) = &H9090902B:    .MC2(15) = &H42DE889
                    .MC2(16) = &H50000000:    .MC2(17) = &HB9909090
                    .MC3(19) = &H90900AE3:    .MC3(20) = &H8D74FF
                    .MC3(21) = &H9090FAE2:    .MC3(22) = &H53FF33FF
                    .MC3(23) = &H90909004:    .MC3(24) = &H2BADC261
                    .MC3(25) = &H3D0853FF:    .MC3(26) = &H1&
                    .MC3(27) = &H23DCE74:     .MC3(28) = &H74000000
                    .MC3(29) = &HAE807:       .MC3(30) = &H90900000
                    .MC3(31) = &H4589C031:    .MC3(32) = &H90DDEBFC
                    .MC3(33) = &HFF0C75FF:    .MC3(34) = &H53FF0475
                    .MC3(35) = &HC310&
                    ' settings within the code
                    .AllocatedDataPtr = lngMCmem
                    .Ordinal = Ordinal
                    .ParamCount = ParamCount
                    PutMem2 VarPtr(.MC3(24)) + 2&, CInt(ParamCount * 4&)
                End With
                ' copy thunk code to executable memory
                RtlMachineCodeCopy ByVal lngMCmem, udtMC, LenB(udtMC)
                ' remember the procedure address (add thunk offset)
                m_TimerProc = lngMCmem + &H14&
                ' now we can initialize the timer
                m_TimerID = SetTimer(0&, 0&, m_Interval, m_TimerProc)
                ' done!
                Exit Sub
            End If
        End If
        ' timer initialization failed for whatever reason, thus timer is disabled
        m_Enabled = False
    End If
End Sub
Private Sub Private_Stop()
    ' only do this if we still have the procedure
    If m_TimerProc Then
        ' kill the timer
        KillTimer 0&, m_TimerID
        ' reset id
        m_TimerID = 0
        ' free the procedure callback
        VirtualFree m_TimerProc, 0&, &H8000& 'MEM_RELEASE
        ' reset procedure pointer to prevent this getting ran twice
        m_TimerProc = 0
    End If
End Sub
Public Sub Reset()
    ' reset counter start value
    m_Start = 0
End Sub
Private Sub Class_Initialize()
    ' initial values: we have it enabled but we have no interval
    m_Enabled = True
    m_Interval = 0
End Sub
Private Sub Class_Terminate()
    Private_Stop
End Sub
' must be the last procedure, ordinal #1!
Private Function TimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    Dim curCounter As Currency
    ' copy Long to Currency
    GetMem4LngToCur tickCount, curCounter
    ' see if start has been initialized (it never will be zero)
    If m_Start > 0 Then
        ' calculate seconds since beginning
        RaiseEvent Timer(curCounter * 10 - m_Start)
    Else
        ' just remember the start time and start from zero
        m_Start = curCounter * 10
        RaiseEvent Timer(0)
    End If
    ' WHATEVER YOU DO, DO NOT CALL Private_Stop FROM HERE!
End Function
