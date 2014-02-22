VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStartTick As Long
Private mCallCount As Long
Private mSecondsElapsed As Double

Public Enum TickCounterErrors
    ErrorBase = 9999
    
End Enum

' Number of milliseconds since Windows was started
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function CalledTooOften(thresholdOfCalls As Long _
                             ) As Boolean
    If mCallCount = thresholdOfCalls Then
        mCallCount = 0
    Else
        CalledTooOften = True
        mCallCount = mCallCount + 1
    End If
End Function


Public Sub Pause(forHowManySeconds As Single)
    On Error GoTo ErrH
    
    Call Sleep(forHowManySeconds * 1000)
    
    On Error GoTo 0
    Exit Sub
ErrH:   ThrowError "Pause" ', "Line: " & Erl
End Sub


Public Property Get SecondIsDivisibleBy(lSecondsDivisor As Long _
                                      ) As Boolean
    
    If (GetTickCount Mod lSecondsDivisor * 1000) = 0 _
        Then SecondIsDivisibleBy = True
    
End Property


Public Property Get CurrentTick() As Long
    CurrentTick = GetTickCount
End Property


Public Function CalledTooSoon(Optional milliSecs As Long = 1000 _
                            ) As Boolean
    CalledTooSoon = GetTickCount() < mStartTick + milliSecs
End Function


Public Function StartTiming() As Long
    On Error GoTo ErrH
    
    mStartTick = GetTickCount           ' store starting tickCount in private var
    StartTiming = mStartTick            ' return stored tickCount
    
    On Error GoTo 0
    Exit Function
ErrH:   ThrowError "StartTiming"
End Function


Public Function StopTiming(Optional formatted As Boolean = True _
                         , Optional formt As String = "#,0.000 sec" _
                         ) As Variant
    On Error GoTo ErrH
    
    mSecondsElapsed = (GetTickCount - mStartTick) / 1000
    
    If formatted Then
        StopTiming = Format$(mSecondsElapsed, formt)
    
    Else
        StopTiming = mSecondsElapsed
    End If
    
    On Error GoTo 0
    Exit Function
ErrH:   ThrowError "StopTiming"
End Function


Public Function Rate(amountOfWork As Variant _
                   , unitOfMeasure As String _
                   , Optional numberFmt As String = "#,#" _
                   ) As String
    Dim workPerSec As Double
    On Error GoTo ErrH
    
    If mSecondsElapsed < 1 Then
        workPerSec = CDbl(amountOfWork)
    Else
        workPerSec = CDbl(amountOfWork) / mSecondsElapsed
    End If
    
    Rate = Format$(workPerSec, numberFmt) & " " & unitOfMeasure
    
    On Error GoTo 0
    Exit Function
ErrH:   ThrowError "Rate" ', "Line: " & Erl
End Function


'Public Function fmtTicks(tickCount As Long _
'                       ) As String
'    Dim tickSec As Double
'    Dim t1 As Double, t2 As Double, t3 As Double
'
'    tickSec = tickCount / 1000
'
'    If tickSec > 60 * 60 * 2 Then          ' if greater than 2 hours...
'        t1 = tickSec \ (60 * 60)
'        t2 = (tickSec - (t1 * 60 * 60)) \ 60
'        t3 = tickSec Mod 60
'        'fmtTicks = CInt(t1) & " hr, " & CInt(t2) & " min, " & CInt(t3) & " sec"
'        fmtTicks = u.plural("hr", CLng(t1), False) _
'          & ", " & u.plural("min", CLng(t2), False) _
'          & ", " & u.plural("sec", CLng(t3), False)
'
'
'    ElseIf tickSec > 100 Then          ' if greater than 100 seconds...
'        t1 = tickSec \ 60
'        t2 = tickSec Mod 60
'        'fmtTicks = CInt(t1) & " min, " & CInt(t2) & " sec"
'        fmtTicks = u.plural("min", CLng(t1), False) _
'          & ", " & u.plural("sec", CLng(t2), False)
'
'    Else
'        fmtTicks = Format$(tickSec, "#,0.00") & " sec"
'    End If
'End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handler Override
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ThrowError(ByVal currentProcedure As String _
                     , Optional ByVal errorDescrption As String _
                     , Optional ByVal errorNumber As TickCounterErrors = -1 _
                     )
    Const MOD_NAME = "TickCounter"
    
    Call ErrorHandler.ThrowError(MOD_NAME _
                               , currentProcedure _
                               , errorDescrption _
                               , errorNumber)
End Sub





