VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "wgtTextCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MODULE_NAME$ = "wgtTextCursor"

Private Enum wgtTextCursorErrors    ' you may make this Public for tests
    ErrorBase = 3000                ' you may adjust this minimum
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As wgtTextCursorErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private WithEvents mBase As cWidgetBase
Attribute mBase.VB_VarHelpID = -1

Private mSymbol As String




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Drawing Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub mBase_Paint(CC As vbRichClient4.cCairoContext _
                      , ByVal xAbs As Single _
                      , ByVal yAbs As Single _
                      , ByVal dx_Aligned As Single _
                      , ByVal dy_Aligned As Single _
                      , UserObj As Object _
                      )
    On Error GoTo Cleanup
    
    With mBase
        
        '  inherit parent's font properties
        '
        With .Parent
            CC.SelectFont .FontName, .FontSize, .ForeColor _
                                   , .FontBold, .FontItalic
        End With
        
        CC.DrawText 0, 0, .ScaleWidth, .ScaleHeight _
                  , mSymbol, False, vbCenter, 0, 1
    End With
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "mBase_Paint" ', "details of error"
End Sub



Public Property Get Symbol() As String
    Symbol = mSymbol
End Property

Public Property Let Symbol(newSymbolChar As String)
    If mSymbol = newSymbolChar Then Exit Property
    mSymbol = newSymbolChar
    Call mBase.Refresh
End Property








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Widget Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get Widget() As cWidgetBase
  Set Widget = mBase
End Property

Public Property Get Widgets() As cWidgets
  Set Widgets = mBase.Widgets
End Property







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo Cleanup
    
    mSymbol = "|"
    
    Set mBase = Cairo.WidgetBase
    With mBase
        .CanGetFocus = False
        
    End With
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    Set mBase = Nothing
    mSymbol = vbNullString
    
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
                  , Optional errorNumbr As wgtTextCursorErrors = -1 _
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
                    , Optional ByVal errorNumbr As wgtTextCursorErrors = -1 _
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
