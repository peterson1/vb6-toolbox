VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSystemConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IConstraint
Const MODULE_NAME$ = "FileSystemConstraints"

Public Enum FileSystemConstraintMethods
    fn_Undefined = -1
    fn_Exists
    fn_DoesNotExist
End Enum
Private mMethod As FileSystemConstraintMethods

Private Enum FileSystemConstraintsErrors    ' you may make this Public for tests
    ErrorBase = 3000                ' you may adjust this minimum
    NotInitted
    UnsupportedMethod
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As FileSystemConstraintsErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Private mPath As String



Private Function IConstraint_Matches(Actual As Variant) As Boolean
    On Error GoTo Cleanup
    
    mPath = CStr(Actual)
    
    Select Case mMethod
        
        Case fn_Exists
            IConstraint_Matches = Found(mPath)
            
        
        Case fn_DoesNotExist
            IConstraint_Matches = Not Found(mPath)
            
        
        Case fn_Undefined:  Err.Raise NotInitted
        Case Else:          Err.Raise UnsupportedMethod
    End Select
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "IConstraint_Matches" ', "details of error"
End Function


Private Sub IConstraint_WriteMessageTo(ByVal Writer As SimplyVBUnit.TextMessageWriter)
    Dim msg$
    On Error GoTo Cleanup
    
    Select Case mMethod
        
        Case fn_Exists
            Writer.WriteLine vbCrLf _
                          & "Expected file/folder does not exist:" _
                          & vbCrLf & """" & mPath & """"
        
        Case fn_DoesNotExist
            Writer.WriteLine vbCrLf _
                          & "File/folder should NOT exist:" _
                          & vbCrLf & """" & mPath & """"
            
        
        Case fn_Undefined:  Err.Raise NotInitted
        Case Else:          Err.Raise UnsupportedMethod
    End Select
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "IConstraint_WriteMessageTo" ', "details of error"
End Sub



Private Sub IConstraint_WriteActualValueTo(ByVal Writer As SimplyVBUnit.TextMessageWriter)

End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As SimplyVBUnit.TextMessageWriter)

End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function Found(fileOrFoldrPath As String _
                     ) As Boolean
    On Error GoTo Cleanup
    
    If FileFound(fileOrFoldrPath) Then
        Found = True
        
    ElseIf FolderFound(fileOrFoldrPath) Then
        Found = True
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Found" ', "details of error"
End Function

Private Function FileFound(pathOfFile As String) As Boolean
    If LenB(pathOfFile) <> 0 Then ' because Dir("") returns something
        On Error Resume Next
            FileFound = LenB(Dir(pathOfFile)) <> 0
        On Error GoTo 0
    End If
End Function

Private Function FolderFound(targetFoldrPath As String) As Boolean
    On Error Resume Next
        FolderFound = GetAttr(targetFoldrPath) And vbDirectory
    On Error GoTo 0
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(constraintMethd As FileSystemConstraintMethods)
    On Error GoTo Cleanup
    
    mMethod = constraintMethd
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Init" ', "details of error"
End Sub









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo Cleanup
    
    mMethod = fn_Undefined
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    'Set mSomeObject = Nothing
    
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
                  , Optional errorNumbr As FileSystemConstraintsErrors = -1 _
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
                    , Optional ByVal errorNumbr As FileSystemConstraintsErrors = -1 _
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
            
            Case UnsupportedMethod
                errorDescrption = "Unsupported constraint method: [" & mMethod & "]."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub
