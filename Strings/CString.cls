VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : Francesco Balena
' * Web Site         : www.geocities.com/ResearchTriangle/6311/
' * Date             : 13/10/98
' * Time             : 10:24
' * Module Name      : class_CString
' * Module Filename  : CString.bas
' * Comment :
' * This class implements the string data type as an
' * object type, and also supports most string functions
' * It offers a few advantages on native VB strings:
' *
' * 1) chars are stored in ANSI format, use less memory
' * 2) you decide how much memory is allocated to each
' *    string, hence memory is not re-allocated with each
' *    new assignment
' * 3) since data is stored as ANSI, no internal conversion
' *    is necessary when printing to file or passing to APIs
' *    (in this case, pass the StrPtr value as a pointer to chars)
' **********************************************************************

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, ByVal fillByte As Integer)

' number of chars initially allocated to the string
Const DEFAULT_MAXLENGTH = 256

' this array holds the character
Private chars() As Byte
' the current lenght of the string
Private m_Len As Long
' the current UBound() of the array (accounts for ending null char)
Private m_MaxLength As Long

Private Sub Class_Initialize()
   ' create the chars() array
   SetBufferSize DEFAULT_MAXLENGTH
End Sub

' set the size of the internal buffer
' Use "mystr.SetBufferSize mystr.Length" to release unused memory

Sub SetBufferSize(ByVal newSize As Long, Optional clearIt As Boolean)
   m_MaxLength = newSize
   If m_Len Or clearIt = 0 Then
      ReDim Preserve chars(0 To m_MaxLength) As Byte
   Else
      ReDim chars(0 To m_MaxLength) As Byte
   End If
End Sub

' the current value as a Unicode string
' this is the default item of the class, so you can use
' the object variable as if it were a string variable

Property Get Value() As String
Attribute Value.VB_UserMemId = 0
   If m_Len > 0 Then
      Value = Space$(m_Len)
      CopyMemory ByVal Value, chars(0), m_Len
   End If
End Property

Property Let Value(NewValue As String)
   ' check that the private array is large enough
   m_Len = Len(NewValue)
   If m_Len > m_MaxLength Then
      SetBufferSize m_Len, True
   End If
   ' copy the characters into the private array
   If m_Len > 0 Then
      CopyMemory chars(0), ByVal NewValue, m_Len
   End If
End Property

' the address of the first character (read-only)

Function StrPtr() As Long
   StrPtr = VarPtr(chars(0))
   ' append a null character to the internal string
   ' in case this will be passed to an external DLL
   chars(m_Len) = 0
End Function

' copy data from a memory address (Friend sub)

Friend Sub CopyData(ByVal sourceAddr As Long, ByVal NumChars As Long)
   ' raise error if new Length is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' check that the private array is large enough
   m_Len = NumChars
   If m_Len > m_MaxLength Then
      SetBufferSize m_Len, True
   End If
   ' copy the characters into the private array
   If m_Len > 0 Then
      CopyMemory chars(0), ByVal sourceAddr, m_Len
   End If
End Sub

' the current Length (can also be written to)
' ("Len" can't be used because it is a reserved word)

Property Get Length() As Long
   Length = m_Len
End Property

Property Let Length(ByVal NewValue As Long)
   ' raise error if new Length is invalid
   If NewValue < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' check that the private array is large enough
   If NewValue > m_MaxLength Then
      SetBufferSize NewValue, False
   End If
   ' fill with blanks if necessary
   If NewValue > m_Len Then
      FillMemory chars(m_Len), NewValue - m_Len, 32
   End If
   ' truncate or expand the string
   m_Len = NewValue
End Property

' clear the string

Sub Clear()
   m_Len = 0
End Sub

' create a copy of this CString object

Function Copy() As CString
   Set Copy = New CString
   Copy.CopyData StrPtr(), m_Len
End Function

' create a new CString object with leftmost characters

Property Get Left(ByVal NumChars As Long) As CString
   ' raise error if new Length is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' adjust length if necessary
   If NumChars > m_Len Then NumChars = m_Len
   ' create a new CString object
   Set Left = New CString
   Left.CopyData StrPtr(), NumChars
End Property

' replace leftmost characters using a given string

Property Let Left(ByVal NumChars As Long, NewValue As Variant)
   ' raise error if new parameter is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' can't copy more character than existing
   If NumChars > m_Len Then NumChars = m_Len
   ' can't copy more characters then the string's length
   If NumChars > Len(NewValue) Then NumChars = Len(NewValue)
   ' copy chars from the string into the local array
   If NumChars > 0 Then
      CopyMemory chars(0), ByVal CStr(NewValue), NumChars
   End If
End Property

' replace leftmost characters using a CString object

Property Set Left(ByVal NumChars As Long, NewValue As CString)
   ' raise error if new parameter is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' can't copy more character than existing
   If NumChars > m_Len Then NumChars = m_Len
   ' can't copy more characters then the string's length
   If NumChars > NewValue.Length Then NumChars = NewValue.Length
   ' copy chars from the string into the local array
   If NumChars > 0 Then
      CopyMemory chars(0), ByVal NewValue.StrPtr, NumChars
   End If
End Property

' create a new CString object with rightmost characters

Property Get Right(ByVal NumChars As Long) As CString
   ' raise error if new Length is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' adjust length if necessary
   If NumChars > m_Len Then NumChars = m_Len
   ' create a new CString object
   Set Right = New CString
   Right.CopyData VarPtr(chars(m_Len - NumChars)), NumChars
End Property

' replace rightmost characters using a given string

Property Let Right(ByVal NumChars As Long, NewValue As Variant)
   ' raise error if new parameter is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' can't copy more character than existing
   If NumChars > m_Len Then NumChars = m_Len
   ' can't copy more characters then the string's length
   If NumChars > Len(NewValue) Then NumChars = Len(NewValue)
   ' copy chars from the string into the local array
   If NumChars > 0 Then
      CopyMemory chars(m_Len - NumChars), ByVal CStr(NewValue), NumChars
   End If
End Property

' replace leftmost characters using a CString object

Property Set Right(ByVal NumChars As Long, NewValue As CString)
   ' raise error if new parameter is invalid
   If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
   ' can't copy more character than existing
   If NumChars > m_Len Then NumChars = m_Len
   ' can't copy more characters then the string's length
   If NumChars > NewValue.Length Then NumChars = NewValue.Length
   ' copy chars from the string into the local array
   If NumChars > 0 Then
      CopyMemory chars(m_Len - NumChars), ByVal NewValue.StrPtr, NumChars
   End If
End Property

' the Mid function
' one- and two-params syntaxes are supported

Property Get Mid(ByVal StartPos As Long, Optional NumChars As Variant) As CString
   Dim bytes As Long
   ' raise error if the first parameter is invalid
   If StartPos < 0 Then Err.Raise 5, TypeName(Me), "Invalid starting position"
   
   If IsMissing(NumChars) Then
      ' if missing, use the remainder of the string
      bytes = m_Len - StartPos + 1
   Else
      bytes = NumChars
      If bytes < 0 Then
         ' raise error if new parameter is invalid
         Err.Raise 5, TypeName(Me), "Invalid length"
      ElseIf StartPos + bytes > m_Len Then
         ' truncate it if too long
         bytes = m_Len - StartPos + 1
      End If
   End If
   
   ' create a new CString object
   Set Mid = New CString
   If StartPos <= m_Len Then
      Mid.CopyData VarPtr(chars(StartPos - 1)), bytes
   End If
End Property

' replace a portion of the string with another string
' if the second parameter is omitted, the string is replaced up to its end

Property Let Mid(ByVal StartPos As Long, Optional NumChars As Variant, NewValue As Variant)
   Dim bytes As Long
   ' raise error if the first parameter is invalid
   If StartPos < 0 Then Err.Raise 5, TypeName(Me), "Invalid starting position"
   
   If IsMissing(NumChars) Then
      ' if missing, use the remainder of the string
      bytes = m_Len - StartPos + 1
   Else
      bytes = NumChars
      If bytes < 0 Then
         ' raise error if new parameter is invalid
         Err.Raise 5, TypeName(Me), "Invalid length"
      ElseIf StartPos + bytes > m_Len Then
         ' truncate it if too long
         bytes = m_Len - StartPos + 1
      End If
   End If
   
   If bytes > Len(NewValue) Then bytes = Len(NewValue)
   If bytes > 0 Then
      CopyMemory chars(StartPos - 1), ByVal CStr(NewValue), bytes
   End If
End Property

' replace a portion of the string with another CString object
' if the second parameter is omitted, the string is replaced up to its end

Property Set Mid(ByVal StartPos As Long, Optional NumChars As Variant, NewValue As CString)
   Dim bytes As Long
   ' raise error if the first parameter is invalid
   If StartPos < 0 Then Err.Raise 5, TypeName(Me), "Invalid starting position"
   
   If IsMissing(NumChars) Then
      ' if missing, use the remainder of the string
      bytes = m_Len - StartPos + 1
   Else
      bytes = NumChars
      If bytes < 0 Then
         ' raise error if new parameter is invalid
         Err.Raise 5, TypeName(Me), "Invalid length"
      ElseIf StartPos + bytes > m_Len Then
         ' truncate it if too long
         bytes = m_Len - StartPos + 1
      End If
   End If
   
   If bytes > NewValue.Length Then bytes = NewValue.Length
   If bytes > 0 Then
      CopyMemory chars(StartPos - 1), ByVal NewValue.StrPtr, bytes
   End If
   
End Property

' the ASC function, but returns -1 if the string is null

Function Asc() As Integer
   If m_Len > 0 Then
      Asc = chars(0)
   Else
      Asc = -1
   End If
End Function

' the ANSI code of a character, or -1 if out-of-range

Property Get Char(ByVal index As Long) As Byte
   If index > 0 Or index <= m_Len Then
      Char = chars(index - 1)
   Else
      Char = -1
   End If
End Property

Property Let Char(ByVal index As Long, ByVal NewValue As Byte)
   ' no effect if index is out of valid range
   If index > 0 Or index <= m_Len Then
      chars(index - 1) = NewValue
   End If
End Property

' find a substring
' only case-sensitive searches are supported

Function InStr(search As Variant, Optional StartPos As Long = 1) As Long
   Dim searchChars() As Byte
   Dim searchLen As Long
   Dim firstChar As Byte
   Dim i As Long, j As Long
   
   If TypeOf search Is CString Then
      ' if a CString object is passed, then make a local copy of it
      searchLen = search.Length
      ' if a null string, exit immediately
      If searchLen = 0 Then InStr = StartPos: Exit Function
      ReDim searchChars(0 To searchLen - 1) As Byte
      CopyMemory searchChars(0), ByVal search.StrPtr, searchLen
   Else
      ' if a string is passed, create the corresponding local array
      searchLen = Len(search)
      ' if a null string, exit immediately
      If searchLen = 0 Then InStr = StartPos: Exit Function
      searchChars() = StrConv(CStr(search), vbFromUnicode)
   End If
   
   ' cache the first character to be searched
   firstChar = searchChars(0)
   
   For i = StartPos - 1 To m_Len - searchLen
      If chars(i) = firstChar Then
         For j = 1 To searchLen - 1
            If chars(i + j) <> searchChars(j) Then Exit For
         Next
         If j = searchLen Then
            InStr = i + 1
            Exit For
         End If
      End If
   Next
   
End Function

' trasform the string to uppercase (returns Me)

Function UCase() As CString
   ' only characters in the range "a-z" are converted
   Dim i As Long, acode As Integer
   For i = 0 To m_Len - 1
      acode = chars(i)
      If acode >= 97 And acode <= 122 Then chars(i) = acode - 32
   Next
   Set UCase = Me
End Function

' trasform the string to lowercase (returns Me)

Function LCase() As CString
   ' only characters in the range "A-Z" are converted
   Dim i As Long, acode As Integer
   For i = 0 To m_Len - 1
      acode = chars(i)
      If acode >= 65 And acode <= 90 Then chars(i) = acode + 32
   Next
   Set LCase = Me
End Function

' reverse the string (returns Me)

Function ReverseStr() As CString
   Dim i As Long, acode As Byte
   For i = 0 To m_Len \ 2 - 1
      acode = chars(i)
      chars(i) = chars(m_Len - 1 - i)
      chars(m_Len - 1 - i) = acode
   Next
   Set ReverseStr = Me
End Function

' append one or more strings or CString objects (returns Me)

' append a string or a CString object (returns Me)

Function Append(arg As Variant) As CString
   Dim i As Long, j As Long
   Dim argObj As CString
   Dim newLength As Long
   Dim argLen As Long
   
   ' prepare the result
   Set Append = Me
   
   If TypeOf arg Is CString Then
      Set argObj = arg
      argLen = argObj.Length
   Else
      argLen = Len(arg)
   End If
   
   ' exit if null string
   If argLen = 0 Then Exit Function
   
   ' prepare the receiving buffer
   newLength = m_Len + argLen
   If newLength > m_MaxLength Then
      SetBufferSize newLength, False
   End If
   
   ' append the string
   If argObj Is Nothing Then
      ' this is a regular string argument
      CopyMemory chars(m_Len), ByVal CStr(arg), argLen
   Else
      ' this is a CString argument
      CopyMemory chars(m_Len), ByVal argObj.StrPtr, argLen
   End If
   
   ' update m_Len to point after the appended characters
   m_Len = newLength
   
End Function

' search a substring using Boyer & Moore algorithm

Function Find(search As String, Optional StartPos As Long = 1) As Long
   Dim bytes() As Byte
   Dim i As Long, searchLen As Long, dist As Long, index As Long
   
   ' save results from previous call
   Static saveSearch As String
   Static distance() As Integer
   
   ' it is better to have the substring in a byte array
   bytes() = StrConv(search, vbFromUnicode)
   searchLen = Len(search)
   index = StartPos + searchLen - 1
   
   ' build the distance table
   ' this is a 256-item array that, for each possible
   ' ANSI character, stores the "distance" of this char
   ' from the end of the search substring (if the char
   ' appears in the substring), or the length of the
   ' substring (if the character doesn't appear in the
   ' substring). Note that the last character in the
   ' substring corresponds to a null value in the table
   
   ' this block is executed only if the search substring
   ' differs from the last call to this method
   If search <> saveSearch Then
      ReDim distance(0 To 255) As Integer
      For i = 0 To 255
         distance(i) = searchLen
      Next
      For i = 1 To searchLen
         distance(bytes(i - 1)) = searchLen - i
      Next
      ' remember for next time
      saveSearch = search
   End If
   
   ' scan the string
   Do While index <= m_Len
      ' retrieve the distance of this character from the
      ' end of the search substring
      dist = distance(chars(index - 1))
      If dist Then
         ' increment Index of found distance - in fact there
         ' is no reason to scan all the characters in the middle
         index = index + dist
      Else
         ' this might be the last character in the string
         ' check if the substring is all there
         For i = 1 To searchLen - 1
            If chars(index - searchLen - 1 + i) <> bytes(i - 1) Then Exit For
         Next
         
         If i = searchLen Then
            ' we've found a match
            Find = index - searchLen + 1
            Exit Function
         End If
         ' the search failed, skip this character and continue
         index = index + 1
      End If
   Loop
   
End Function

' write the string to a binary file
' optionally prefix it with string length

Sub WriteToFile(ByVal FileNum As Integer, Optional PrefixLength As Boolean)
   If PrefixLength Then
      Put #FileNum, , m_Len
   End If
   ' temporarily shorten the internal buffer
   ' (we need this to use VB's Put statement)
   ReDim Preserve chars(0 To m_Len) As Byte
   ' write all chars to file
   Put #FileNum, , chars()
   ' restore the buffer
   ReDim Preserve chars(0 To m_MaxLength) As Byte
End Sub

' read the string from a binary file
' if PrefixLength is True, uses stored length prefix, else
' uses current string length

Sub ReadFromFile(ByVal FileNum As Integer, Optional PrefixLength As Boolean)
   If PrefixLength Then
      ' read the string length
      Get #FileNum, , m_Len
      ' enlarge the buffer if necessary
      If m_Len > m_MaxLength Then
         SetBufferSize m_Len, True
      End If
   End If
   ' temporarily shorten the internal buffer
   ' (we need this to use VB's Put statement)
   ReDim Preserve chars(0 To m_Len) As Byte
   ' get chars from file
   Get #FileNum, , chars()
   ' restore the buffer
   ReDim Preserve chars(0 To m_MaxLength) As Byte
End Sub

' truncate the string to its first null char, returns new length

' this function is useful when dealing with API calls that return
' null-terminate strings in the buffer

Function TrimToNull() As Long
   Dim i As Long
   ' assume there is no null char
   m_Len = m_MaxLength
   ' scan the string searching for the first null character
   For i = 1 To m_MaxLength
      If chars(i) = 0 Then
         ' set new length and exit
         m_Len = i - 1
         Exit For
      End If
   Next
   ' new length is also the return value
   TrimToNull = m_Len
End Function

