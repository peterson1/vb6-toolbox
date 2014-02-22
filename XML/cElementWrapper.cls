VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cElementWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum cElementWrapperErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ItemNotFound
    NoChildren
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As cElementWrapperErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Dim mElement As cElement






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Simple Getters
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get Element() As cElement
    Set Element = mElement
End Property

Public Property Get Name() As String
    Name = mElement.TagName
End Property

Public Property Get ChildCount() As Long
    ChildCount = mElement.ChildElements.Count
End Property

Public Property Get Children() As cElements
    Set Children = mElement.ChildElements
End Property

'Public Function NewEnum() As IUnknown
'   Set NewEnum = mElement.ChildElements '.[_NewEnum]
'End Function




Public Property Get Text() As String
    On Error GoTo Cleanup
    With mElement
    
        ' find end of opening tag
        Dim i&: i = InStr(.XML, ">") + 1
        
        ' find start of closing tag
        Dim j&: j = InStrRev(.XML, "<") - 1
        
        ' if found at 1st char,
        '  - tag is self-closing
        '  - return blank
        If j = 0 Then Exit Property
        
        ' return text between those positions
        Text = Trim$(Mid$(.XML, i, j - i + 1))
        
    End With
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
            
            If mError.HasError Then
                'MsgBox ""
            End If
            
  LoadError "Text" ', "details of error"
End Property


Public Property Get HasTag(nameOfTag As String) As Boolean
    HasTag = Me.TagIndex(nameOfTag) <> -1
End Property


Public Property Get TagIndex(nameOfTag As String) As Long
    Dim e As cElement
    
    For Each e In mElement.ChildElements
        
        If LCase$(e.TagName) _
         = LCase$(nameOfTag) Then Exit Function
        
        TagIndex = TagIndex + 1
    Next e
    
    
    '  return -1 of none found
    '
    TagIndex = -1
End Property


Public Function Tag(tagNameOrIndx As Variant _
                  , Optional errorIfNoneFound As Boolean = True _
                  ) As cElementWrapper
    Dim i&
    On Error GoTo Cleanup
    
    If Not mElement.HasChildElements Then
        ErrorIf errorIfNoneFound _
             , "Parent element has no children." _
              , cElementWrapperErrors.NoChildren
        Exit Function
    End If
    
    If IsNumeric(tagNameOrIndx) Then
        i = CLng(tagNameOrIndx)
    
    Else
        i = Me.TagIndex(CStr(tagNameOrIndx))
        If i = -1 Then
            ErrorIf errorIfNoneFound _
                 , "Tag not found:  <" & tagNameOrIndx & ">." _
                  , cElementWrapperErrors.ItemNotFound
            Exit Function
        End If
    End If
    
    Set Tag = New cElementWrapper
    Call Tag.Init(mElement.ChildElements(i))
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Tag" ', "details of error"
End Function


Public Function Find(searchFiltr As String _
                   , Optional errorIfNoneFound As Boolean = True _
                   ) As cElementWrapper
    Dim attName$, attVal, pos&, e As cElement, a As cAttribute
    On Error GoTo Cleanup
    
    If Not mElement.HasChildElements Then
        ErrorIf errorIfNoneFound _
             , "Parent element has no children." _
              , cElementWrapperErrors.NoChildren
        Exit Function
    End If
              
    pos = InStr(searchFiltr, "=")
    attName = Trim$(Left$(searchFiltr, pos - 1))
    attVal = Trim$(Right$(searchFiltr, Len(searchFiltr) - pos))
    
    For Each e In mElement.ChildElements
        For Each a In e.Attributes
            If LCase$(a.Name) = LCase$(attName) Then
                If LCase$(a.Text) = LCase$(attVal) Then
                    
                    Set Find = New cElementWrapper
                    Call Find.Init(e)
                    Exit Function
                    
                End If
            End If
        Next a
    Next e
    
    ErrorIf errorIfNoneFound _
         , "Tag not found where:  " & searchFiltr _
          , cElementWrapperErrors.ItemNotFound
    
Cleanup:    SaveError
            Set a = Nothing
            Set e = Nothing
            'Erase ss
  LoadError "Find" ', "details of error"
End Function


Public Property Get A_(attributeNameOrIndex As Variant _
                     , Optional valueIfNoSuchAttribute As Variant _
                     ) As Variant
Attribute A_.VB_UserMemId = 0
    Dim i&
    On Error GoTo Cleanup
    With mElement.Attributes
        
        '  if no attributes,
        '   - either raise an error -OR-
        '   - return default value (if given)
        '
        If .Count = 0 Then
            ErrorIf IsMissing(valueIfNoSuchAttribute) _
                 , "Tag has no attributes."
            A_ = valueIfNoSuchAttribute:    Exit Property
        End If
        
        
        '  if arg is numeric,
        '   - use it as index
        '
        If IsNumeric(attributeNameOrIndex) Then
            i = CLng(attributeNameOrIndex)
            
            '  if out of bounds,
            '   - either raise an error -OR-
            '   - return default value (if given)
            '
            If i > .Count - 1 Then
                ErrorIf IsMissing(valueIfNoSuchAttribute) _
                     , "Attribute index out of bounds:  [" & i & "]."
                A_ = valueIfNoSuchAttribute:    Exit Property
            End If
        
        
        '  otherwise,
        '   - get index of attribute that matches the name
        Else
            i = Me.AttributeIndex(CStr(attributeNameOrIndex))
            
            '  if no such attribute,
            '   - either raise an error -OR-
            '   - return default value (if given)
            '
            If i = -1 Then
                ErrorIf IsMissing(valueIfNoSuchAttribute) _
                     , "Attribute not found: """ & attributeNameOrIndex & """."
                A_ = valueIfNoSuchAttribute:    Exit Property
            End If
        End If
        
        A_ = .Item(i).Text
        
    End With
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Get]A_" ', "details of error"
End Property


Public Property Get AttributeIndex(attrbuteName As String) As Long
    Dim a As cAttribute
    
    For Each a In mElement.Attributes
        
        If LCase$(a.Name) _
         = LCase$(attrbuteName) Then Exit Function
        
        AttributeIndex = AttributeIndex + 1
    Next a
    
    
    '  return -1 of none found
    '
    AttributeIndex = -1
End Property


Public Property Get HasAttribute(attrbuteName As String) As Boolean
    HasAttribute = Me.AttributeIndex(attrbuteName) <> -1
End Property


Public Property Get Self() As cElementWrapper
    Set Self = Me
End Property




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(cElementObj As cElement)
    Set mElement = cElementObj
End Sub

Public Sub InitAny(cElementObj As Variant)
    Set mElement = cElementObj
End Sub









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo Cleanup
    
    'Set mSomeObject = New Something
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    Set mElement = Nothing
    
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
                  , Optional errorNumbr As cElementWrapperErrors = -1 _
                  )
    If errCondition Then Err.Raise errorNumbr, TypeName(Me), errorMsg
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
                    , Optional ByVal errorNumbr As cElementWrapperErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = vbCrLf & TypeName(Me) & "." & currntProcedure & "()" _
                        & vbCrLf & "Tag name:  <" & mElement.TagName & ">"
        
        If errorNumbr = -1 Then errorNumbr = .Number
        
        Select Case errorNumbr
            Case NotInitted
                errorDescrption = TypeName(Me) & " not initted." & vbCrLf _
                               & "Please call " & TypeName(Me) _
                               & ".Init() before " & currntProcedure & "."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub

