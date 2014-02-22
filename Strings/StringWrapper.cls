VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  based on Francesco Balena's CString class
'
Option Explicit
Const MODULE_NAME$ = "StringWrapper"
Const DEFAULT_MAXLENGTH& = 256

Private Enum StringWrapperErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    InvalidLength
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As StringWrapperErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, ByVal fillByte As Integer)



' the current lenght of the string
Private mLength&

' this array holds the character
Private mBytes() As Byte

' the current UBound() of the array (accounts for ending null char)
Private mMaxLength As Long


' current value as a Unicode string
'
Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    On Error GoTo Cleanup
    
    If mLength > 0 Then
        Value = Space$(mLength)
        CopyMemory ByVal Value, mBytes(0), mLength
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Value" ', "details of error"
End Property

Property Let Value(strText As String)
    Call Me.Init(strText)
End Property



Property Get Length() As Long
   Length = mLength
End Property

Property Let Length(ByVal strLength As Long)
    On Error GoTo Cleanup
    
    ErrorIf strLength < 0, "Invalid Length", InvalidLength
    
    ' check that the private array is large enough
    If strLength > mMaxLength Then _
        SetBufferSize strLength, False
    
    ' fill with blanks if necessary
    If strLength > mLength Then _
        FillMemory mBytes(mLength), strLength - mLength, 32
    
    ' truncate or expand the string
    mLength = strLength
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Length" ', "details of error"
End Property


' the address of the first character (read-only)
'
Public Function StrPtr() As Long
   
   StrPtr = VarPtr(mBytes(0))
   
   ' append a null character to the internal string
   ' in case this will be passed to an external DLL
   mBytes(mLength) = 0
End Function




Public Property Get Has(findThisStr As String _
                      , Optional caseSensitiv As Boolean = True _
                      , Optional startPos As Long = 1 _
                      ) As Long
    On Error GoTo Cleanup
    
    If caseSensitiv Then
        'Has = InStr(startPos, mText, findThisStr, vbBinaryCompare)
        'Has = InStr2(findThisStr, startPos)
        Has = Find(findThisStr, startPos)
        
    Else
        Has = InStr(startPos, Me.Value, findThisStr, vbTextCompare)
    End If
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Has" ', "details of error"
End Property





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(strText As String)
    On Error GoTo Cleanup
    
    ' check that the private array is large enough
    mLength = Len(strText)
    If mLength > mMaxLength Then SetBufferSize mLength, True
    
    ' copy the characters into the private array
    If mLength > 0 Then
        CopyMemory mBytes(0), ByVal strText, mLength
        'mBytes() = StrConv(strText, vbFromUnicode)
    End If
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Init" ', "details of error"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function Find(findThisStr As Variant _
                    , Optional startPos As Long = 1 _
                    ) As Long
   Dim byts() As Byte, i&, searchLen&, dist&, indx&
    On Error GoTo Cleanup
    
   ' save results from previous call
   Static saveSrch$
   Static dstance%()
   
   ' it is better to have the substring in a byte array
   byts() = StrConv(findThisStr, vbFromUnicode)
   searchLen = Len(findThisStr)
   indx = startPos + searchLen - 1
   
   ' build the dstance table
   ' this is a 256-item array that, for each possible
   ' ANSI character, stores the "dstance" of this char
   ' from the end of the findThisStr substring (if the char
   ' appears in the substring), or the length of the
   ' substring (if the character doesn't appear in the
   ' substring). Note that the last character in the
   ' substring corresponds to a null value in the table
   
   ' this block is executed only if the findThisStr substring
   ' differs from the last call to this method
   If findThisStr <> saveSrch Then
      ReDim dstance(0 To 255) As Integer
      For i = 0 To 255
         dstance(i) = searchLen
      Next
      For i = 1 To searchLen
         dstance(byts(i - 1)) = searchLen - i
      Next
      ' remember for next time
      saveSrch = findThisStr
   End If
   
   ' scan the string
   Do While indx <= mLength
      ' retrieve the dstance of this character from the
      ' end of the findThisStr substring
      dist = dstance(mBytes(indx - 1))
      If dist Then
         ' increment Index of found dstance - in fact there
         ' is no reason to scan all the characters in the middle
         indx = indx + dist
      Else
         ' this might be the last character in the string
         ' check if the substring is all there
         For i = 1 To searchLen - 1
            If mBytes(indx - searchLen - 1 + i) <> byts(i - 1) Then Exit For
         Next
         
         If i = searchLen Then
            ' we've found a match
            Find = indx - searchLen + 1
            Exit Function
         End If
         ' the findThisStr failed, skip this character and continue
         indx = indx + 1
      End If
   Loop
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase byts
  LoadError "Find" ', "details of error"
End Function


'
'  doesn't work
'
' only case-sensitive searches are supported
Private Function InStr2(findThisStr As Variant _
                      , Optional startPos As Long = 1 _
                      ) As Long
    Dim searchChars() As Byte, searchLen&, firstChar As Byte, i&, j&
    On Error GoTo Cleanup
    
    If TypeOf findThisStr Is StringWrapper Then
    
        ' if a StringWrapper object is passed, then make a local copy of it
        searchLen = findThisStr.Length
        
        ' if a null string, exit immediately
        If searchLen = 0 Then InStr2 = startPos: Exit Function
        
        ReDim searchChars(0 To searchLen - 1) As Byte
        CopyMemory searchChars(0), ByVal findThisStr.StrPtr, searchLen
    
    
    Else
        ' if a string is passed, create the corresponding local array
        searchLen = Len(findThisStr)
        
        ' if a null string, exit immediately
        If searchLen = 0 Then InStr2 = startPos: Exit Function
        
        searchChars() = StrConv(CStr(findThisStr), vbFromUnicode)
    End If
    
    
    ' cache the first character to be searched
    firstChar = searchChars(0)
    
    For i = startPos - 1 To mLength - searchLen
        If mBytes(i) = firstChar Then
            For j = 1 To searchLen - 1
                If mBytes(i + j) <> searchChars(j) Then Exit For
            Next
            
            If j = searchLen Then
                InStr2 = i + 1
                Exit For
            End If
            
        End If
    Next
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase searchChars
  LoadError "InStr2" ', "details of error"
End Function



' set the size of the internal buffer
' Use "mystr.SetBufferSize mystr.Length" to release unused memory

Private Sub SetBufferSize(ByVal buffrSize As Long _
                        , Optional clearBuffr As Boolean _
                        )
    On Error GoTo Cleanup
    
    mMaxLength = buffrSize
    
    If mLength Or clearBuffr = 0 Then
        ReDim Preserve mBytes(0 To mMaxLength) As Byte
    Else
        ReDim mBytes(0 To mMaxLength) As Byte
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "SetBufferSize" ', "details of error"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo Cleanup
    
   ' create the chars() array
   SetBufferSize DEFAULT_MAXLENGTH
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    Erase mBytes
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Terminate"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As StringWrapperErrors = -1 _
                  )
    If errCondition Then Err.Raise errorNumbr, MODULE_NAME, errorMsg
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
                    , Optional ByVal errorNumbr As StringWrapperErrors = -1 _
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

