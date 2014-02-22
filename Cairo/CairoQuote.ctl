VERSION 5.00
Begin VB.UserControl CairoQuote 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CairoQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Private Enum CairoQuoteErrors    ' you may make this Public for tests
'    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
'    NotInitted
'    AlreadyInitted
'    ' add error numbers here
'End Enum
'
'Private Type ErrorHolder            '
'    HasError As Boolean             '  temp storage for errors
'    Source As String                '
'    Number As CairoQuoteErrors    '
'    Description As String
'End Type
'Private mError As ErrorHolder
'
'
'
'
'Public Sub DrawQuote()
'    Debug.Print "DrawQuote"
'End Sub
'
'
'
'
'
'' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
''
''    Constructor
''
'' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'Public Sub Init()
'
'End Sub
'
'
'
'
'
'
'
'
'
'' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
''
''    Class Events
''
'' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'Private Sub UserControl_Initialize()
'    On Error GoTo Cleanup
'
'    'Set mSomeObject = New Something
'
'Cleanup:    SaveError
'            'Set someObj = Nothing
'            'Erase someArray
'  LoadError "UserControl_Initialize"
'End Sub
'
'Private Sub UserControl_Terminate()
'    On Error GoTo Cleanup
'
'    'Set mSomeObject = Nothing
'
'Cleanup:    SaveError
'            'Set someObj = Nothing
'            'Erase someArray
'  LoadError "UserControl_Terminate"
'End Sub
'
'
'
'
'
'' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
''
''    Error Handlers
''
'' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'Private Sub ErrorIf(errCondition As Boolean _
'                  , errorMsg As String _
'                  , Optional errorNumbr As CairoQuoteErrors = -1 _
'                  )
'    If errCondition Then Err.Raise errorNumbr, TypeName(Me), errorMsg
'End Sub
'
'Private Sub SaveError()
'    With mError
'        If Err Then
'            .HasError = True
'            .Description = Err.Description
'            .Number = Err.Number
'            .Source = Err.Source
'
'        Else
'            .HasError = False
'            .Description = vbNullString
'            .Number = 0
'            .Source = vbNullString
'        End If
'    End With
'    Err.Clear
'End Sub
'
'Private Sub Blame(ByVal currntProcedure As String _
'                , Optional ByVal errorDescrption As String _
'                , Optional ByVal errorNumbr As CairoQuoteErrors = -1 _
'                )
'    Call SaveError
'    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
'End Sub
'
'Private Sub LoadError(ByVal currntProcedure As String _
'                    , Optional ByVal errorDescrption As String _
'                    , Optional ByVal errorNumbr As CairoQuoteErrors = -1 _
'                    )
'    With mError
'        If Not .HasError Then Exit Sub
'
'        If LenB(errorDescrption) = 0 Then
'            errorDescrption = .Description
'        Else
'            errorDescrption = .Description & vbCrLf & errorDescrption
'        End If
'
'        currntProcedure = TypeName(Me) & "." & currntProcedure & "()"
'
'        If errorNumbr = -1 Then errorNumbr = .Number
'
'        Select Case errorNumbr
'            Case NotInitted
'                errorDescrption = TypeName(Me) & " not initted." & vbCrLf _
'                               & "Please call " & TypeName(Me) _
'                               & ".Init() before " & currntProcedure & "."
'
'            Case Else
'                errorDescrption = currntProcedure & vbCrLf & errorDescrption
'        End Select
'
'        Err.Raise errorNumbr, .Source, errorDescrption
'
'    End With
'End Sub
