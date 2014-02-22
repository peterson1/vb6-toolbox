VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "wgtTextField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MODULE_NAME$ = "wgtTextField"
Const CARET_KEY$ = "CursorWidgetKey"

Const FONT_NAME$ = "Verdana"
Const FONT_SIZE! = 8

Const SINGLE_LINE As Boolean = True
Const H_ALIGN& = vbLeftJustify

Const PADDING_LEFT& = 10

Private Enum wgtTextFieldErrors     ' you may make this Public for tests
    ErrorBase = 3000                ' you may adjust this minimum
                                    ' add more errors here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As wgtTextFieldErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private WithEvents mBase As cWidgetBase
Attribute mBase.VB_VarHelpID = -1

Private mText As cStringBuilder _
      , mCaret As wgtTextCursor





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    TextField Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get Text() As String
    Text = mText.ToString
End Property

Public Property Let Text(newTextVal As String)
    On Error GoTo Cleanup
    
    With mText
        If .ToString = newTextVal Then Exit Property
        Call .Clear
        Call .Append(newTextVal)
    End With
    
    Call mBase.Refresh
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Text" ', "details of error"
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
'    Tab Focus Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub mBase_GotFocus()
    On Error GoTo Cleanup
    With mBase
        .BackColor = RGB(254, 255, 225)
        .BorderColor = vbBlue
        
        mCaret.Widget.Visible = True
        
    End With
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "mBase_GotFocus" ', "details of error"
End Sub


Private Sub mBase_LostFocus()
    On Error GoTo Cleanup
    With mBase
        .BackColor = vbWhite
        .BorderColor = RGB(140, 140, 140)
        
        mCaret.Widget.Visible = False
        
    End With
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "mBase_LostFocus" ', "details of error"
End Sub






Private Sub mBase_KeyPress(KeyAscii As Integer)
    On Error GoTo Cleanup
    
    Select Case KeyAscii
        
        Case vbKeyBack
            Me.Text = TrimRight(mText.ToString, 1)
            
        Case Else
            Call mText.Append(Chr(KeyAscii))
            Call mBase.Refresh
        
    End Select
        
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "mBase_KeyPress" ', "details of error"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Drawing Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub mBase_Paint(CC As cCairoContext _
                      , ByVal xAbs As Single _
                      , ByVal yAbs As Single _
                      , ByVal DX As Single _
                      , ByVal DY As Single _
                      , UserObj As Object _
                      )
    On Error GoTo Cleanup
    With mBase
        
        '  set boundaries
        '
        CC.SetLineWidth 1
        CC.Rectangle 0, 0, .Width, .Height, True   '<-- note the last optional Param, which ensures a Pixel-aligned drawing inside the Widgets-Bounds
        
        
        '  fill background
        '
        If .BackColor <> -1 Then 'a .BackColor of -1 would be our indicator for no BackGround-Fill at all...
            CC.SetSourceColor .BackColor '... so we enter here (ensuring a Fill), only if we got a real BackColor
            CC.Fill True  'as already mentioned a few times in the Cairo-Tutorial-examples, the True-Param ensures waiting for the final Stroke
        End If
        
        
        '  draw the border
        '
        CC.SetSourceColor .BorderColor
        CC.Stroke
        
        
        '  draw text
        '
        CC.SelectFont FONT_NAME, FONT_SIZE
        CC.DrawText PADDING_LEFT, 0, .ScaleWidth, .ScaleHeight _
                  , mText.ToString, SINGLE_LINE _
                  , H_ALIGN, 0, 1      'and the H- and V-Centered Text-Drawing
        
        
        '  update caret
        '
        'mCaret.Text = mText.ToString
        
    End With
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "mBase_Paint" ', "details of error"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get TrimRight(strText As String _
                            , charsToTrim As Long _
                            ) As String
    Dim haba&
    On Error GoTo Cleanup
    
    haba = Len(strText)
    
    If haba = 0 Then                    '  if zero-length,
        Exit Property                   '   - end here (returns blank)
        
    ElseIf haba = charsToTrim Then      '  if all will be trimmed,
        Exit Property                   '   - end here (returns blank)
        
    ElseIf haba < charsToTrim Then      '  if too short,
        TrimRight = strText             '   - return orig text
    
    Else
        TrimRight = Left$(strText, haba - charsToTrim)
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "TrimRight" ', "details of error"
End Property











' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo Cleanup
    
    Set mText = New_RC4.StringBuilder
    
    Set mBase = Cairo.WidgetBase
    mBase.CanGetFocus = True
    
    Set mCaret = mBase.Widgets.Add(New wgtTextCursor, CARET_KEY)
    
    Call mBase_LostFocus
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    Set mCaret = Nothing
    Set mBase = Nothing
    Set mText = Nothing

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
                  , Optional errorNumbr As wgtTextFieldErrors = -1 _
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
                    , Optional ByVal errorNumbr As wgtTextFieldErrors = -1 _
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
