VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IConstraint

Public Enum XmlConstraintMethods
    fn_Undefined = -1
    fn_NameIs
    fn_HasTag
    fn_ValueIs
    fn_TextIs
    
End Enum

Private Enum XmlConstraintsErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    UnsupportedMethod
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As XmlConstraintsErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mMethod As XmlConstraintMethods, mExpected$
Private mWrapper As cElementWrapper, mParentTag$, mAttributeKey$, mSearchFilter$




Private Function IConstraint_Matches(Actual As Variant) As Boolean
    On Error GoTo Cleanup
    
    Set mWrapper = Nothing
    Set mWrapper = New cElementWrapper
    With mWrapper
        Call .InitAny(Actual)
        
        '  if parent tag given...
        '
        If LenB(mParentTag) <> 0 Then
            .Init .Tag(mParentTag).Element
            mParentTag = vbNullString
        End If
        
        If LenB(mSearchFilter) <> 0 Then
            .Init .Find(mSearchFilter).Element
            mSearchFilter = vbNullString
        End If
        
        
        Select Case mMethod
            
            
            Case fn_NameIs
                IConstraint_Matches = LCase$(.Name) _
                                    = LCase$(mExpected)
                
                
            Case fn_HasTag
                IConstraint_Matches = .HasTag(mExpected)
            
            
            
            Case fn_ValueIs
                IConstraint_Matches = LCase$(.A_(mAttributeKey)) _
                                    = LCase$(mExpected)
            
            
            Case fn_TextIs
                IConstraint_Matches = .Text _
                                    = mExpected
            
                
            Case fn_Undefined:  Err.Raise NotInitted
            Case Else:          Err.Raise UnsupportedMethod
        End Select
    End With
Cleanup:    SaveError
            'Set e = Nothing
            'Erase someArray
  LoadError "IConstraint_Matches" ', "details of error"
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As SimplyVBUnit.TextMessageWriter)

End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As SimplyVBUnit.TextMessageWriter)

End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As SimplyVBUnit.TextMessageWriter)
    Dim msg$
    On Error GoTo Cleanup
    Select Case mMethod
        
        Case fn_NameIs
            Writer.WriteLine _
                vbCrLf & " Expected tag name was:  <" & mExpected & ">." _
              & vbCrLf & "but Actual tag name is:  <" & mWrapper.Name & ">."
            
            
        Case fn_HasTag
            Writer.WriteLine _
                vbCrLf & "Tag <" & mWrapper.Name & "> does not contain tag <" & mExpected & ">."
        
        
        Case fn_ValueIs
            Writer.WriteLine _
                vbCrLf & "Tag name:  <" & mWrapper.Name & ">" _
              & vbCrLf & " Expected value of attribute [" & mAttributeKey & "] was:  """ & mExpected & """." _
              & vbCrLf & "   but Actual value is:  """ & mWrapper(mAttributeKey) & """."
        
        
        Case fn_TextIs
            Writer.WriteLine _
                vbCrLf & "Tag name:  <" & mWrapper.Name & ">" _
              & vbCrLf & " Expected text was:  """ & mExpected & """." _
              & vbCrLf & "but Actual text is:  """ & mWrapper.Text & """."
        
        
        Case fn_Undefined:  Err.Raise NotInitted
        Case Else:          Err.Raise UnsupportedMethod
    End Select
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "IConstraint_WriteMessageTo" ', "details of error"
End Sub



Public Function A_(attrbuteKey As String) As XmlConstraints
Attribute A_.VB_UserMemId = 0
    mAttributeKey = attrbuteKey
    Set A_ = Me
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Direct Aliases
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Function NameIs(nameOfTag As String) As IConstraint
    Set NameIs = Tag_.NameIs(nameOfTag)
End Function

Public Function HasTag(nameOfTag As String) As IConstraint
    Set HasTag = Tag_.HasTag(nameOfTag)
End Function

Public Function ValueIs(attrbuteValue As Variant) As IConstraint
    Set ValueIs = Tag_.ValueIs(attrbuteValue)
End Function

Public Function TextIs(textOfTag As String) As IConstraint
    Set TextIs = Tag_.TextIs(textOfTag)
End Function

Public Function Find(searchFiltr As String) As XmlConstraints
    Set Find = Tag_.Find(searchFiltr)
End Function







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(constraintMethd As XmlConstraintMethods _
              , expectedVal As Variant)
    On Error GoTo Cleanup
    
    mMethod = constraintMethd
    mExpected = CStr(expectedVal)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Init" ', "details of error"
End Sub

Public Property Let ParentTag(nameOfParent As String)
    mParentTag = nameOfParent
End Property

Public Property Let SearchFilter(searchFiltr As String)
    mSearchFilter = searchFiltr
End Property







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
    
    Set mWrapper = Nothing
    
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
                  , Optional errorNumbr As XmlConstraintsErrors = -1 _
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
                    , Optional ByVal errorNumbr As XmlConstraintsErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = TypeName(Me) & "." & currntProcedure & "()"

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
