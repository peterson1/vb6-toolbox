VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'! StringBuilder class is similar to other SringBuilder classes available by
'! programming language libraries such as java.  It is intended to simplify
'! long string concatenation sequences and to simplify padding strings with
'! quotes or other markup.
'!
'! It is also anywhere from 30 - 200+ times faster than concatenating strings using the '&' or '+'
'! operators. Since it makes appending strings a linear operation, the longer the string being
'! appended, the better the run-time improvement.
'! @see http://support.microsoft.com/kb/170964
'! @see http://roymacleanvba.wordpress.com/2010/01/14/performance-strings/
'!
'! @author Dave Parillo, Rite Solutions, Inc.
'! $Date: 2010-01-14 08:32:34 -0800 (Thu, 14 Jan 2010) $
'! $Revision: 703 $
'!


Private Const pInitial As Long = 32     '* Initial Length of the string buffer

Private pTotal As Long                  '* Total Length of the buffer
Private pCurrent As Long                '* Length of the string value within the buffer
Private pBuf As String                  '* The string buffer

'* Definitions for various ways to return strings surrounded by other text.
'* Typically quotation marks, but other styles are used also.
'*
'*  @see Quotify
Enum sbQuoteStyle
    sbNone = 0
    sbDouble = 1
    sbSingle = 2
    sbSquareBrackets = 3
    sbExcelDateString = 4
    sbADODBSheetReference = 5
End Enum

'* Set up the string buffer.
'* It's defined to have an initial buffer size,
'* but the 'string size' (current length) of the buffer is 0
Private Sub Class_Initialize()
    pCurrent = 0
    pTotal = pInitial
    pBuf = Space(pTotal)
End Sub

'* Tear down the string buffer.
Private Sub Class_Terminate()
    pBuf = vbNullString
End Sub

'* Sets or resets the string to a value.
Private Sub SetString(ByRef StringValue As String)
    pBuf = Space(pTotal)
    pCurrent = 0
    Me.Append (StringValue)
End Sub

'* Returns the current StringBuilder string.
'*
'* Value is the default StringBuilder class member. Given:
'* <code>
'* Dim foo As StringBuilder
'* Set foo = New StringBuilder
'* foo.Value = "Hello, world!"
'* </code>
'* Then:
'* - foo.Value
'* - foo
'*
'* both refer to the current value of foo.
Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = Left$(pBuf, pCurrent)
End Property

'* Sets the string to a value.
'* @param StringValue The string property is set to this value.
Property Let Value(ByRef StringValue As String)
    SetString StringValue
End Property

'* Get the raw string buffer string, including all extra space in the buffer.
'* Used primarily for testing, you won't typically refer to the buffer.
'* Use StringBuilder::Value instead.
Property Get Buffer() As String
    Buffer = pBuf
End Property

'* Returns the current StringBuilder string, optionally surrounded by
'* some type of markup.
'*
'* Given:
'* <code>
'* Dim foo As StringBuilder
'* Set foo = New StringBuilder
'* foo = "Hello, world!"
'* </code>
'* Then:
'* - foo.Value
'* - foo
'* - foo.ToString
'* - foo.ToString(sbNone)
'*
'* All refer to the current value of foo.
'* @param QuotationType surrounds the string with quotation marks
'*                      (or some other character) if desired.
'*     - the StringBuilder string is not modified when specifying
'*       a QuotationType.
'* @return the current value of the StringBuilder string
'*         surrounded by the specified sbQuoteStyle
Property Get ToString(Optional ByVal QuotationType As sbQuoteStyle) As String
    ToString = Quotify(Me.Value, QuotationType)
End Property

'* Request the length of the StringBuilder string.
'* @return Current length of the StringBuilder string.
Property Get Length() As Long
    Length = pCurrent
End Property

'* Request the length of the StringBuilder raw string buffer.
'* Used primarily for testing, you won't typically refer to the buffer.
'* Use StringBuilder::Length instead.
Property Get BufferLength() As Long
    Debug.Assert pTotal = Len(pBuf)
    BufferLength = pTotal
End Property

'* Reinitalizes the string buffer.
Sub Clear(Optional ByVal TotalLength As Long)
    pCurrent = 0
    If TotalLength = 0 Then
        pTotal = pInitial
    Else
        pTotal = TotalLength
    End If
    pBuf = Space(pTotal)
End Sub

'* Attempts to reduce storage used for the string.
'*
'* Removes both leading and trailing blank spaces from pBuf and sets
'* the buffer lengths to match.
Sub TrimToSize()
    pBuf = Trim$(pBuf)
    pCurrent = Len(pBuf)
    pTotal = pCurrent
End Sub

'* Appends to the end of the string property.
'* @param theString The string that is appended to the string property.
'*        Passed by Reference only for speed.
'* @param QuoteType the Type of Quotation Style to apply to the StringBuilder string
Sub Append(ByRef theString As String, _
           Optional ByVal QuoteType As sbQuoteStyle)
    Dim s As String
    Dim bufLen As Long
    Dim sLen As Long
    
    s = Quotify(theString, QuoteType)
    sLen = Len(s)
    bufLen = pCurrent + sLen  'length of new string

    If bufLen <= pTotal Then
        'MUCH faster than pBuf = pBuf & s
        Mid(pBuf, pCurrent + 1, sLen) = s
    Else
        ' Increase buffer size
        While pTotal < bufLen
          pTotal = pTotal * 2
        Wend
        pBuf = Left$(pBuf, pCurrent) _
             & s _
             & Space$(pTotal - bufLen)
    End If
    pCurrent = bufLen
End Sub

'* Inserts a string at a specific index of the string property
'* @param theString The string being inserted.
'* @param theIndex The index location where the string is to be inserted
'* @param QuoteType the Type of Quotation Style to apply to the StringBuilder string
Sub Insert(ByRef theString As String, _
           ByVal theIndex As Integer, _
           Optional ByVal QuoteType As sbQuoteStyle)
    Dim s As String
    Dim lhs As String
    Dim rhs As String
    Dim sLen As Long
    
    s = Quotify(theString, QuoteType)
    sLen = Me.Length + Len(s)
    lhs = Left$(pBuf, theIndex)
    rhs = Mid$(pBuf, theIndex + 1, Me.Length - theIndex)
    SetString lhs
    Me.Append s
    Me.Append rhs
End Sub

'* Places mark up around a string.
'* @note  theStr is not modified by Quotify.
'* @param QuoteType the Type of Quotation Style to apply to the StringBuilder string
'* @return the current value of the theStr string
'*         surrounded by the specified sbQuoteStyle
Private Function Quotify(ByRef theStr As String, _
                         ByVal QuoteType As sbQuoteStyle) As String
    Dim leftChar As String
    Dim rightChar As String
    If QuoteType = sbNone Then
        Quotify = theStr
    Else
        Select Case QuoteType
            Case sbDouble
                leftChar = Chr(34)
                rightChar = leftChar
            Case sbSingle
                leftChar = Chr(39)
                rightChar = leftChar
            Case sbSquareBrackets
                leftChar = "["
                rightChar = "]"
            Case sbExcelDateString
                leftChar = "#"
                rightChar = leftChar
            Case sbADODBSheetReference
                leftChar = " ["
                rightChar = "$] "
            Case Else
                leftChar = ""
                rightChar = leftChar
        End Select
        Quotify = leftChar & theStr & rightChar
    End If
End Function

