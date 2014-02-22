VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum XmlWriterErrors           ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As XmlWriterErrors       '
    Description As String
End Type
Private mError As ErrorHolder

Public Parent As XmlWriter

Private mTagName$, mInnerText$

Private mAttributeCount&, mAttributeKeys$(), mAttributeValues() As Variant

Private mTagCount&, mTags() As XmlWriter


Public Function Level() As Long
    Dim parnt As XmlWriter
    On Error GoTo Cleanup
    
    Set parnt = Me.Parent
    
    Do Until parnt Is Nothing
        Set parnt = parnt.Parent
        Level = Level + 1
    Loop
    
Cleanup:    SaveError
            Set parnt = Nothing
            'Erase someArray
  LoadError "Level" ', "details of error"
End Function


Public Property Get ToString(Optional indentd As Boolean = False _
                           ) As String
    Dim innrText$
    On Error GoTo Cleanup
    
    innrText = Me.InnerText(indentd)
    
    
    ' Windows ANSI
    '<?xml version="1.0" encoding="ISO-8859-1"?>
    
    
    '  if no inner text/tag,
    '   - write self-closing tag
    '   - indent as needed
    '
    If LenB(innrText) = 0 Then
        ToString = "<" & mTagName _
                       & Me.Attributes() & "/>"
        
        If Me.Level <> 0 And indentd Then _
            ToString = vbCrLf & String$(Me.Level, vbTab) & ToString
    
    
    '  otherwise,
    '   - write usual <open> and </close> tags
    '
    Else
        ToString = "<" & mTagName _
                       & Me.Attributes() & ">"
        
        ' indent open tag and inner text as needed
        If Me.Level <> 0 And indentd Then _
            ToString = vbCrLf & String$(Me.Level, vbTab) _
                              & ToString _
                     & vbCrLf & String$(Me.Level + 1, vbTab)
        
        
        ToString = ToString & innrText
        
        ' indent close tag as needed
        If Me.Level <> 0 And indentd Then _
            ToString = ToString _
                     & vbCrLf & String$(Me.Level, vbTab)
        
        ToString = ToString & "</" & mTagName & ">"
    End If
    
    
    

Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
            innrText = vbNullString
  LoadError "ToString" ', "details of error"
End Property


Public Function ToFile(pathOfFile As String _
                     , Optional indentd As Boolean = False _
                     ) As String
    On Error GoTo Cleanup
    
    Dim fn&: fn = FreeFile
    Open pathOfFile For Output As #fn

    Print #fn, Me.ToString(indentd);

    Close #fn
    
    '  return path
    ToFile = pathOfFile
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "ToFile" ', "details of error"
End Function


Public Property Get Attributes(Optional pairsDelimitr$ = " " _
                             , Optional strQuotes$ = """" _
                             ) As String
    Dim i&, ss$()
    On Error GoTo Cleanup
    
    If mAttributeCount = 0 Then Exit Property
    
    ReDim ss(mAttributeCount - 1)
    For i = 0 To UBound(ss)
        ss(i) = mAttributeKeys(i) _
              & "=" _
              & strQuotes _
              & mAttributeValues(i) _
              & strQuotes
    Next i
    
    Attributes = " " & Join(ss, pairsDelimitr) & " "
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "[Get]Attributes" ', "details of error"
End Property


Public Property Let Attributes(Optional pairsDelimitr$ = " " _
                             , Optional strQuotes$ = """" _
                             , joinedAttrbutes As String)
    On Error GoTo Cleanup
    
    If LenB(joinedAttrbutes) = 0 Then Exit Property
    
    Dim ss$(): ss = SplitPairs(joinedAttrbutes, "=" _
                             , strQuotes, pairsDelimitr)
    
    Dim i&: For i = 0 To UBound(ss, 1)
        Me.A_(ss(i, 0)) = ss(i, 1)
    Next i
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "[Let]Attributes" ', "details of error"
End Property


Public Function SplitPairs(ByVal joindPairs As String _
                         , Optional keyValSeparatr$ = "=" _
                         , Optional valQuotes$ = """" _
                         , Optional pairsDelimitr$ = " " _
                         ) As String()
    Dim a1$(), a2$(), i&, j&, c$, atKey$, atVal$, atCount&
    On Error GoTo Cleanup
    joindPairs = Trim$(joindPairs)
    
    i = 1 ' start with first character (position 1)
    Do
        
        ' find end of attribute key
        j = InStr(i, joindPairs, keyValSeparatr) - 1
        
        atKey = Trim$(Mid$(joindPairs, i, j - i + 1))
        
        
        ' find start of attribute value
        If LenB(valQuotes) = 0 Then
            i = j + Len(keyValSeparatr) + 1
        Else
            i = InStr(j + Len(keyValSeparatr), joindPairs, valQuotes) + Len(valQuotes)
        End If
        
        ' find end of attribute value
        j = InStr(i, joindPairs, valQuotes & pairsDelimitr) - 1
        
        ' if end reached, subtract length of quotes
        If j = -1 Then _
            j = Len(joindPairs) - Len(valQuotes)
        
        atVal = Mid$(joindPairs, i, j - i + 1)
        
        ReDim Preserve a1(1, atCount)   '
        a1(0, atCount) = atKey          '  add key & value to array
        a1(1, atCount) = atVal          '
        atCount = atCount + 1           '
        
        
        ' start with character after pairs delimiter
        i = j + Len(valQuotes & pairsDelimitr) + 1
        
    Loop Until i >= Len(joindPairs)
    
    
    
    '  return transposed array a1
    '
    ReDim a2(UBound(a1, 2), 1)
    For i = 0 To UBound(a1, 2)
        a2(i, 0) = a1(0, i)
        a2(i, 1) = a1(1, i)
    Next i
    
    SplitPairs = a2
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase a1, a2
  LoadError "SplitPairs" ', "details of error"
End Function


Public Property Let InnerText(Optional indentd As Boolean = False _
                            , strText As String)
    mInnerText = strText
End Property


Public Property Get InnerText(Optional indentd As Boolean = False _
                            ) As String
    Dim i&, ss$(), delimitr$
    On Error GoTo Cleanup
    
    If mTagCount = 0 Then
        InnerText = mInnerText
'        If indentd Then InnerText _
'            = vbCrLf & String$(Me.Level, vbTab) & InnerText
    
    Else
        ReDim ss(mTagCount - 1)
        For i = 0 To UBound(ss)
            ss(i) = mTags(i).ToString(indentd)
        Next i
        
        If indentd Then _
            delimitr = vbCrLf & String$(Me.Level, vbTab)
        
        InnerText = Join(ss, delimitr) _
                  & delimitr _
                  & mInnerText
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "[Get]InnerText" ', "details of error"
End Property



Public Property Let A_(attrbuteKey As Variant _
                     , attrbuteValue As Variant _
                     )
    On Error GoTo Cleanup
    
    '  see if the key already exists
    '
    Dim i&: i = Me.AttributeIndex(attrbuteKey)
    
    
    '  if no such key, store new key & value
    '   - otherwise, replace existing value
    '
    If i = -1 Then
        If UBound(mAttributeKeys) < mAttributeCount _
            Then Call ExpandAttributesArray
        
        i = mAttributeCount ' use next position
        mAttributeKeys(i) = CStr(attrbuteKey)
        mAttributeCount = mAttributeCount + 1
    End If
    
    mAttributeValues(i) = attrbuteValue
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Let]A_" ', "details of error"
End Property

Public Property Get A_(attrbuteKey As Variant _
                     ) As Variant
Attribute A_.VB_UserMemId = 0
    Dim indx&
    On Error GoTo Cleanup
    
    indx = Me.AttributeIndex(attrbuteKey)
    
    ErrorIf indx = -1, "Attribute not found: """ & attrbuteKey & """"
    
    A_ = mAttributeValues(indx)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Get]A_" ', "details of error"
End Property


Public Property Get Tag(ByVal nameOfTag As String _
                      ) As XmlWriter
    On Error GoTo Cleanup
    
    ErrorIf LenB(nameOfTag) = 0 _
         , "Name of tag cannot be blank."
    
    If UBound(mTags) < mTagCount Then ExpandTagsArray
    
    Set mTags(mTagCount) = New XmlWriter
    With mTags(mTagCount)
        .Name = nameOfTag
    End With
    
    Set Tag = mTags(mTagCount)
    
    Set Tag.Parent = Me
    
    mTagCount = mTagCount + 1
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Get]Tag" ', "details of error"
End Property

Public Property Set Tag(ByVal nameOfTag As String _
                      , xmlWritrObj As XmlWriter)
    On Error GoTo Cleanup
    
    '  if blank name given,
    '   - use name of param object
    '
    If LenB(nameOfTag) = 0 Then _
        nameOfTag = xmlWritrObj.Name
    
    With Me.Tag(nameOfTag)
        Let .Attributes = xmlWritrObj.Attributes
        Let .InnerText = xmlWritrObj.InnerText
    End With
    
Cleanup:    SaveError
            Set xmlWritrObj = Nothing
            'Erase someArray
  LoadError "[Set]Tag" ', "details of error"
End Property



Public Property Get Encode(strText$) As String
    Dim i&, c$, ss$()
    On Error GoTo Cleanup
    
    If LenB(strText) = 0 Then Exit Property
    
    ReDim ss(Len(strText) - 1)
    
    For i = 0 To UBound(ss)
        c = Mid$(strText, i + 1, 1)
        
        If XmlSafe(c) Then
            ss(i) = c
        Else
            ss(i) = EntityEncode(c)
        End If
    Next i
    
    Encode = Join(ss)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "Encode" ', "details of error"
End Property


Public Property Get XmlSafe(charactr$) As Boolean
    Dim a%: a = Asc(charactr)
    
    XmlSafe = True
    
    If a > 62 And a < 127 Then Exit Property     ' ? @ A-Z [ \ ] ^ _ ` a-z { | } ~

    If a = 32 Then Exit Property                ' space
    If a > 39 And a < 60 Then Exit Property     ' ( ) * + , - . /  0-9 : ;
    If a = 33 Then Exit Property                ' !
    If a > 34 And a < 38 Then Exit Property     ' # $ %
    If a = 61 Then Exit Property                ' =
    
    
    XmlSafe = False
End Property


Private Property Get EntityEncode(singleCharactr$) As String
    Dim a%
    On Error GoTo Cleanup
    
    a = Asc(singleCharactr)
    
    Select Case a

        Case 9:     EntityEncode = "&vb_Tab;"   ' tab
        Case 10:    EntityEncode = "&vb_Lf;"    ' line feed
        Case 13:    EntityEncode = "&vb_Cr;"    ' carriage return
        Case 34:    EntityEncode = "&quot;"     ' double-quote
        Case 38:    EntityEncode = "&amp;"      ' ampersand
        Case 39:    EntityEncode = "&apos;"     ' apostrophe
        Case 60:    EntityEncode = "&lt;"       ' less-than sign
        Case 62:    EntityEncode = "&gt;"       ' greater-than sign
        Case 133:   EntityEncode = "&hellip;"   ' horizontal ellipsis (= three dot leader)
        Case 146:   EntityEncode = "&rsquo;"    ' right single quote ’
        Case 147:   EntityEncode = "&ldquo;"    ' left double quote “
        Case 148:   EntityEncode = "&rdquo;"    ' right double quote ”
        Case 150:   EntityEncode = "&ndash;"    ' en dash –
        Case 151:   EntityEncode = "&mdash;"    ' em dash —
        Case 225:   EntityEncode = "&aacute;"   ' á : Latin small letter a with acute accent
        Case 237:   EntityEncode = "&iacute;"   ' í : Latin small letter i with acute accent
        
                
        Case Else
            Err.Raise -1, , "No entity encoding for character """ _
                           & singleCharactr & """ : Asc(" & a & ")."
    End Select
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "EntityEncode" ', "details of error"
End Property





Public Property Get Decode(strText$) As String
    On Error GoTo ErrH
    Decode = strText
    
    ReplaceIn Decode, "&vb_Tab;", Chr(9)
    ReplaceIn Decode, "&vb_Lf;", Chr(10)
    ReplaceIn Decode, "&vb_Cr;", Chr(13)
    ReplaceIn Decode, "&quot;", """"
    ReplaceIn Decode, "&amp;", "&"
    ReplaceIn Decode, "&apos;", "'"
    ReplaceIn Decode, "&lt;", "<"
    ReplaceIn Decode, "&gt;", ">"
    ReplaceIn Decode, "&hellip;", "…"
    ReplaceIn Decode, "&rsquo;", "’"
    ReplaceIn Decode, "&ldquo;", "“"
    ReplaceIn Decode, "&rdquo;", "”"
    ReplaceIn Decode, "&ndash;", "–"
    ReplaceIn Decode, "&mdash;", "—"
    ReplaceIn Decode, "&aacute;", "á"
    ReplaceIn Decode, "&iacute;", "í"
    
ErrH: Blame "Decode"
End Property







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Idiot-Proof Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get Self() As XmlWriter
    Set Self = Me
End Property

Public Property Let Name(nameOfTag As String)
    mTagName = nameOfTag
End Property

Public Property Get Name() As String
    Name = mTagName
End Property

Public Property Get AttributeCount() As Long
    AttributeCount = mAttributeCount
End Property

Public Property Get AttributeIndex(attrbuteKey As Variant) As Long
    If mAttributeCount = 0 Then
        AttributeIndex = -1
    Else
        AttributeIndex = IndexOf(CStr(attrbuteKey), mAttributeKeys)
    End If
End Property

Public Property Get AttributeExists(attrbuteKey As Variant) As Boolean
    AttributeExists = Me.AttributeIndex(attrbuteKey) <> -1
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ReplaceIn(ByRef strText As String _
                    , lookForThisStr As String _
                    , replaceWithStr As String _
                    , Optional compareMethd As VbCompareMethod = vbTextCompare _
                    )
    On Error GoTo ErrH
    
    '  if not found
    '   - do not proceed
    '
    If InStr(1, strText, lookForThisStr, compareMethd) = 0 Then Exit Sub
    
    strText = Replace(strText, lookForThisStr _
                             , replaceWithStr _
                    , Compare:=compareMethd)
    
ErrH: Blame "ReplaceIn"
End Sub

Private Function IndexOf(lookForStr As String _
                       , lookIn1D() As String _
                       ) As Long
    
    Dim i&: For i = 0 To UBound(lookIn1D)
        
        '  if found, stop looking
        '
        If LCase$(lookForStr) _
         = LCase$(lookIn1D(i)) Then Exit Function
         
        IndexOf = IndexOf + 1
    Next i
    
    '  return -1 of none found
    '
    IndexOf = -1
End Function



Private Sub ExpandAttributesArray()
    On Error GoTo Cleanup
    
    Const EXPANSION_INCREMENT& = 4
    
    If mAttributeCount = 0 Then
        ReDim mAttributeKeys(EXPANSION_INCREMENT)
        ReDim mAttributeValues(EXPANSION_INCREMENT)
    Else
        ReDim Preserve mAttributeKeys(mAttributeCount + EXPANSION_INCREMENT)
        ReDim Preserve mAttributeValues(mAttributeCount + EXPANSION_INCREMENT)
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "ExpandAttributesArray" ', "details of error"
End Sub

Private Sub ExpandTagsArray()
    On Error GoTo Cleanup
    
    Const EXPANSION_INCREMENT& = 8
    
    If mTagCount = 0 Then
        ReDim mTags(EXPANSION_INCREMENT)
    Else
        ReDim Preserve mTags(mTagCount + EXPANSION_INCREMENT)
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "ExpandTagsArray" ', "details of error"
End Sub












' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =










' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo Cleanup
    
    Call ExpandAttributesArray
    
    Call ExpandTagsArray
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    Erase mAttributeKeys, mAttributeValues
    
    Dim i&: For i = 0 To mTagCount - 1
        Set mTags(i) = Nothing
    Next i
    Erase mTags
    
    Set Me.Parent = Nothing
    
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
                  , Optional errorNumbr As XmlWriterErrors = -1 _
                  )
    If errCondition Then Err.Raise errorNumbr, TypeName(Me), errorMsg
End Sub

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As XmlWriterErrors = -1 _
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
                    , Optional ByVal errorNumbr As XmlWriterErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = vbCrLf & TypeName(Me) & "." & currntProcedure & "()" _
                        & vbCrLf & "Tag name:  <" & mTagName & ">"

        If errorNumbr = -1 Then errorNumbr = .Number
        
        Select Case errorNumbr
            Case NotInitted
                errorDescrption = TypeName(Me) & " not initted." & vbCrLf _
                               & "Please call " & TypeName(Me) _
                               & ".ExpandAttributesArray() before " & currntProcedure & "."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub

