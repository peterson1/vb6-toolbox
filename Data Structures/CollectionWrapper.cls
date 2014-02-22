VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum CollectionWrapperErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As CollectionWrapperErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mCollection As VBA.Collection

Private mTypeName$



Public Function Add(objToAdd As Variant) As Variant
    On Error GoTo ErrH
    
    If mTypeName = vbNullString _
        Then mTypeName = TypeName(objToAdd)
    
    Call mCollection.Add(objToAdd)
    
    If IsObject(objToAdd) Then
        Set Add = objToAdd
    Else
        Add = objToAdd
    End If
    
ErrH: Blame "Add"
End Function


Public Sub Clear()
    On Error GoTo ErrH
    With mCollection
    
        Dim i&: For i = 1 To .Count
            Call .Remove(1)
        Next i
        
    End With
ErrH: Blame "Clear"
End Sub


Public Property Get Count() As Long
    Count = mCollection.Count
End Property


Public Property Get Item(zeroBasedIndx As Long) As Object
    On Error GoTo ErrH
    
    ErrorIf Me.Count = 0, "Collection has no items."
    
    ErrorIf zeroBasedIndx >= Me.Count _
         , "Attempted to retrieve item at index [" & zeroBasedIndx & "]." & vbCrLf _
         & "But collection is only up to index [" & Me.Count - 1 & "]."
    
    Set Item = mCollection.Item(zeroBasedIndx + 1)
    
ErrH: Blame "[Get]Item"
End Property


Public Sub Remove(zeroBasedIndx As Long)
    On Error GoTo ErrH
    
    Call mCollection.Remove(zeroBasedIndx + 1)
    
ErrH: Blame "Remove"
End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCollection.[_NewEnum]
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo ErrH
    
    Set mCollection = New VBA.Collection
    
ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Call Me.Clear
    
    Set mCollection = Nothing
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As CollectionWrapperErrors = -1 _
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

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As CollectionWrapperErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As CollectionWrapperErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = mTypeName & "s" _
                        & "." & currntProcedure & "()"

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

