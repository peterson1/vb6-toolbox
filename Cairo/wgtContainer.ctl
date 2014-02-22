VERSION 5.00
Begin VB.UserControl wgtContainer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "wgtContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MODULE_NAME$ = "wgtContainer"

Private Enum wgtContainerErrors        ' you may make this Public for tests
    ErrorBase = 3000                    ' you may adjust this minimum
                                        ' add more errors here
End Enum

Private Type ErrorHolder                '
    HasError As Boolean                 '  temp storage for errors
    Source As String                    '
    Number As wgtContainerErrors     '
    Description As String
End Type
Private mError As ErrorHolder




Public Property Get WidgetRoot() As cWidgetRoot
    Static mWidgetRoot As cWidgetRoot
    On Error GoTo Cleanup
    
    If mWidgetRoot Is Nothing Then
        'Set mWidgetRoot = New cWidgetRoot
        Set mWidgetRoot = Cairo.WidgetRoot
        
        mWidgetRoot.RenderContentIn Me
    End If
    
    Set WidgetRoot = mWidgetRoot
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Get]WidgetRoot" ', "details of error"
End Property

Public Property Get Widgets() As cWidgets
    Set Widgets = WidgetRoot.Widgets
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Control Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub UserControl_Initialize()
    On Error GoTo Cleanup
    
    '  initialize vbRichClient4
    '   - needed for using Cairo
    '
    Call RC4Factory.Init(App.Path)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "UserControl_Initialize" ', "details of error"
End Sub








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As wgtContainerErrors = -1 _
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
                    , Optional ByVal errorNumbr As wgtContainerErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = MODULE_NAME & "." & currntProcedure & "()"

        errorDescrption = currntProcedure & vbCrLf & errorDescrption
        
        If errorNumbr = -1 Then errorNumbr = .Number

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub
