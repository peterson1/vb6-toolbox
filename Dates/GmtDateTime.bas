Attribute VB_Name = "GmtDateTime"
Option Explicit

' derived from:
'   http://www.vbrad.com/article.aspx?id=62

Private Enum GmtDateTimeErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As GmtDateTimeErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer

    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer

    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer

    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long

End Type

Private Declare Function GetTimeZoneInformation _
    Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
Private Const TIME_ZONE_ID_STANDARD& = 1
Private Const TIME_ZONE_ID_UNKNOWN& = 0

Private Const TIME_ZONE_ID_DAYLIGHT& = 2


Private Declare Function InternetTimeToSystemTime Lib "wininet.dll" _
       (ByVal lpszTime As String, _
        ByRef pst As SYSTEMTIME, _
        ByVal dwReserved As Long) _
        As Long



'  Figures out GMT time based on StartingDate.
'
'   - will use Now() if no starting date
'
Public Function GmtDate(Optional ByVal startngDate As Variant _
                      ) As Date
    On Error GoTo ErrH
    
    Dim diff&: diff = GmtTimeDifference
    
    If IsMissing(startngDate) Then startngDate = Now
    
    GmtDate = DateAdd("s", -diff, startngDate)
    
ErrH: Blame "GmtDate"
End Function



'  Returns time difference between
'    local & GMT time in seconds.
'
'  If the  result is negative,
'    your time zone lags behind GMT zone.
'
'  If the  result is positive,
'    your time zone is ahead.
'
Public Function GmtTimeDifference() As Long
    Dim tz As TIME_ZONE_INFORMATION, retCode&
    
    
    ' retrieve the time zone information from API
    retCode = GetTimeZoneInformation(tz)
    
    
    ' convert to seconds
    GmtTimeDifference = -tz.Bias * 60
    
    
    ' if we are in daylight  saving time, apply the bias.
    If retCode = TIME_ZONE_ID_DAYLIGHT& Then
        
        ' if tz.DaylightDate.wMonth = 0 then the daylight saving time change doesn't occur
        If tz.DaylightDate.wMonth <> 0 Then GmtTimeDifference _
            = GmtTimeDifference - tz.DaylightBias * 60
    End If
    
ErrH: Blame "GmtTimeDifference"
End Function



'  Returns the time in your local time zone
'    which corresposponds to GMT time.
'
'  param: gmtRefernceDate
'   Provides the time & date from which to make calculations
'
Public Function GmtDateLocal(gmtRefernceDate As Date) As Date
    On Error GoTo ErrH
    
    GmtDateLocal = DateAdd("s", GmtTimeDifference(), gmtRefernceDate)
    
ErrH: Blame "GmtDateLocal"
End Function




'  Converts internet time format to VB Date.
'
'   supported formats:  Rfc822 and Iso8601
'
'     Iso8601 is either
'           1997-07-16T19:20:30+01:00 (25 bytes)
'        or 1997-07-16T19:20:30Z      (20 bytes)
'
'     Rfc822 is
'           Tue, 23 Sep 2003 13:21:00 -07:00 (32 bytes)
'        or Tue, 23 Sep 2003 13:21:00 GMT (29 bytes)
'
'     The key difference is that Iso8661 time has a latin letter T in position 11
'
Public Function GmtStringToDate(ByVal internetDateStr As String _
                              ) As Date
    On Error GoTo ErrH
    
    internetDateStr = Trim$(internetDateStr)
    
    If Mid$(DateString, 11, 1) = "T" Then
        GmtStringToDate = Iso8601TimeToLocalVbTime(internetDateStr)
    Else
        GmtStringToDate = Rfc822TimeToLocalVbTime(internetDateStr)
    End If
    
ErrH: Blame "GmtStringToDate"
End Function


Public Function GmtFormat(dateObj As Date _
                        , Optional dateFormt$ = "ddd, dd mmm yyyy hh:nn:ss" _
                        , Optional gmtSuffx$ = " GMT" _
                        ) As String
    On Error GoTo ErrH
    
    GmtFormat = Format$(dateObj, dateFormt) & gmtSuffx
    
ErrH: Blame "GmtFormat"
End Function



Private Function Iso8601TimeToLocalVbTime(ByVal sIso8601 As String) As Date

    'format of the time is similar to this: 1997-07-16T19:20:30+01:00
    'or                                     1997-07-16T19:20:30Z or 2003-10-09T09:40:46Z
    'where Z is UTC (aka GMT time)
    
    'formatting breakdown
    '                1012141618202224
    '       1997-07-16T19:20:30+01:00

    '       1234567891113151719212325
    
    Dim sYear As String
    Dim sMonth As String

    Dim sDay As String
    Dim sHour As String
    Dim sMinute As String

    Dim sSecond As String
    Dim sTimeZone As String
    Dim dtDateTime As Date

    Dim bSign As Boolean
    Dim dGMT As Long

    sYear = Left$(sIso8601, 4)
    sMonth = Mid$(sIso8601, 6, 2)
    sDay = Mid$(sIso8601, 9, 2)
    sHour = Mid$(sIso8601, 12, 2)
    sMinute = Mid$(sIso8601, 15, 2)
    sSecond = Mid$(sIso8601, 18, 2)
    sTimeZone = Mid$(sIso8601, 20)
    
    dtDateTime = CDate(DateSerial(sYear, sMonth, sDay) & " " & _
                                     TimeSerial(sHour, sMinute, sSecond))
                                     
                                         
    'replace Z with +00:00 for easier processing
    sTimeZone = Replace(sTimeZone, "Z", "+00:00", , , vbTextCompare)
    
    'get the size
    bSign = IIf(Left$(sTimeZone, 1) = "+", True, False)
    

    'grab the hour & minutes
    dGMT = Val(Mid$(sTimeZone, Len(sTimeZone) - 3, 2)) + (CInt(Right$(sTimeZone, 2)) * 100 / 60)
    
    If bSign Then
        dtDateTime = DateAdd("H", -dGMT, dtDateTime)
    Else

        dtDateTime = DateAdd("H", dGMT, dtDateTime)
    End If
    
    Iso8601TimeToLocalVbTime = GmtDateLocal(dtDateTime)
    
    
    
End Function

Private Function Rfc822TimeToLocalVbTime(sRfc822 As String) As Date

    Dim uSystemTime As SYSTEMTIME
    Dim sWWW As String
    Dim iHours As Integer

    Dim dGMT As Long
    Dim sHourDifferential As String
    Dim dtDateTime As Date

    
    Dim sSign As String
    Dim bSign As Boolean
    Dim sEscapedTime As String

    Dim sTimeZoneString As String
    Dim iPos As Integer
    'true = positive

    'false = negative
    
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    sWWW = sRfc822
    
    If InStr(1, sWWW, "GMT", vbTextCompare) > 0 Then
        sWWW = Replace(sWWW, "GMT", "+0000")
    End If

    'check to make sure that the time zone is included

    If Len(Trim$(sWWW)) = 25 Then
        'add time zone
        sWWW = sWWW & " +0000"
    End If

    
    If (InStr(1, sWWW, ",") = 0) Then sWWW = "Thu, " & sWWW
    Call InternetTimeToSystemTime(sWWW, uSystemTime, 0&)
    With uSystemTime
        dtDateTime = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & _
                                     TimeSerial(.wHour, .wMinute, .wSecond))
    End With

    
    'get the sign from the back end
    'remove colons, in case the time is 07:00 instead of 0700
    sEscapedTime = Replace(sWWW, ":", "")
    
    sSign = Mid$(sEscapedTime, Len(sEscapedTime) - 4, 1)
    bSign = IIf(sSign = "-", False, True)
    
    'grab the hour & minutes

    iPos = InStrRev(sWWW, " ")
    If iPos > 0 Then
        'get rid of the space and the +/- sign
        sTimeZoneString = Mid$(sWWW, iPos + 2)
        
        'escape it
        sTimeZoneString = Replace(sTimeZoneString, ":", "")
        sTimeZoneString = Replace(sTimeZoneString, " ", "")
        

        'at this point we should have the following: 0700
        dGMT = Val(Left$(sTimeZoneString, 2)) + Val(Right$(sWWW, 2)) * 100 / 60
        
        'dGMT = Val(Mid$(sWWW, Len(sWWW) - 3, 2)) + (CInt(Right$(sWWW, 2)) * 100 / 60)
    Else
        dGMT = 0
    End If

        
        
    If bSign Then
        dtDateTime = DateAdd("H", -dGMT, dtDateTime)
    Else
        dtDateTime = DateAdd("H", dGMT, dtDateTime)
    End If
    
    Rfc822TimeToLocalVbTime = GmtDateLocal(dtDateTime)
    

End Function






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As GmtDateTimeErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As GmtDateTimeErrors = -1 _
                  )
    If errCondition Then ErrRaise errorMsg, errorNumbr
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

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As GmtDateTimeErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As GmtDateTimeErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = "GmtDateTime." & currntProcedure & "()"

        If errorNumbr = -1 Then errorNumbr = .Number
        
        Select Case errorNumbr
            Case NotInitted
                errorDescrption = "GmtDateTime not initted." & vbCrLf _
                               & "Please call GmtDateTime" _
                               & ".Init() before " & currntProcedure & "."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        ErrRaise errorDescrption, errorNumbr, .Source
            
    End With
End Sub

