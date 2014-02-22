VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuotePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum QuotePresenterErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As QuotePresenterErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private WithEvents mText As VB.TextBox _
      , WithEvents mCaption As VB.Label _
      , WithEvents mScroll As VB.HScrollBar _
      , WithEvents mRandom As VB.CommandButton
Attribute mText.VB_VarHelpID = -1
Attribute mCaption.VB_VarHelpID = -1
Attribute mScroll.VB_VarHelpID = -1
Attribute mRandom.VB_VarHelpID = -1

Private mPanel As MSComctlLib.Panel _
      , WithEvents mPanelUpdate As SelfTimer
Attribute mPanelUpdate.VB_VarHelpID = -1

Private mInitted As Boolean


Public Sub Resize()
    On Error GoTo ErrH
    
    If Not mInitted Then Exit Sub
    
    With mScroll
        .Left = .Container.Width - .Width - 100
        .Top = 200
        mRandom.Top = .Top
        mRandom.Left = .Left + 240
    End With
    
    With mText
        '.BorderStyle = 1
        
        .Width = .Container.Width / 2
        .Left = .Width - (.Width / 2)
        
        .Height = TextBoxHeight(mText, .Text)
        .Top = .Container.Height / 2 - .Height / 2 - 600
    End With
    
    With mCaption
        .Top = mText.Top + mText.Height + 150
        .Left = mText.Left
        .Width = mText.Width
    End With
    
ErrH: Blame "Resize"
End Sub

Private Function TextBoxHeight(textBoxCtrl As VB.TextBox _
                             , strText As String _
                             ) As Double
    Dim numRows&, c As cCairoContext
    On Error GoTo Cleanup
    
    
    With textBoxCtrl
        
        Set c = New_RC4.Cairo.CreateSurface(.Width, .Height).CreateContext
        
        With .Font
            c.SelectFont .Name, .Size, _
                       , .Bold, .Italic _
                       , .Underline, .Strikethrough
        End With
        
        numRows = c.DrawText(0, 0 _
                , (.Width - 0) / Screen.TwipsPerPixelX _
                , 0, strText)
        
        TextBoxHeight = numRows * c.GetFontHeight()
        
        TextBoxHeight = TextBoxHeight * Screen.TwipsPerPixelY + 100
        
    End With
    
Cleanup:    SaveError
            Set c = Nothing
            'Erase someArray
  LoadError "TextBoxHeight" ', "details of error"
End Function



Private Sub mPanelUpdate_Timer(ByVal Seconds As Currency)
    On Error GoTo ErrH
    With mPanel
        
        .Text = "  Blessed Mother, guide my work.  "
        
        Select Case .Alignment
            Case sbrCenter:     .Alignment = sbrRight
            Case sbrLeft:       .Alignment = sbrCenter
            Case sbrRight:      .Alignment = sbrLeft
        End Select
    
    End With
ErrH: Blame "mPanelUpdate_Timer"
End Sub



Private Sub mRandom_Click()
    On Error GoTo ErrH
    With mScroll
    
        .Value = Rand.mNumber(.Min, .Max)
        
    End With
ErrH: Blame "mRandom_Click"
End Sub

Private Sub mScroll_Change()
    On Error GoTo ErrH
    
    Call DrawQuote
    
ErrH: Blame "mScroll_Change"
End Sub




Private Sub DrawQuote()
    Dim quot$, capt$
    On Error GoTo ErrH
    
    quot = Rand.mPearl(capt, mScroll.Value)
    
    ' double the line breaks
    mText.Text = Replace(quot, vbCrLf, vbCrLf & vbCrLf)
    
    mCaption.Caption = "-- " & capt
    
    Call Me.Resize
    
ErrH: Blame "DrawQuote"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(txtBoxForQuoteText As VB.TextBox _
              , labelForQuoteCaption As VB.Label _
              , scrollBarForQuoteNav As VB.HScrollBar _
              , cmdBtnForRandom As VB.CommandButton _
              , scrolBarPanelForBottomQuotes As MSComctlLib.Panel _
              )
    On Error GoTo ErrH
    
    Set mText = txtBoxForQuoteText
    Set mCaption = labelForQuoteCaption
    Set mScroll = scrollBarForQuoteNav
    Set mRandom = cmdBtnForRandom
    Set mPanel = scrolBarPanelForBottomQuotes
    
    Set mPanelUpdate = New SelfTimer
    mPanelUpdate.Interval = 1000! * 60 * 5 'five minutes
    
    With mScroll
        .Min = 1
        .Max = PEARLS_MAX
    End With
    
    mInitted = True
    
    Call mRandom_Click
    
ErrH: Blame "Init"
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
    
    Set mText = Nothing
    Set mCaption = Nothing
    Set mScroll = Nothing
    Set mRandom = Nothing
    Set mPanel = Nothing
    Set mPanelUpdate = Nothing
    
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
                  , Optional errorNumbr As QuotePresenterErrors = -1 _
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
                , Optional ByVal errorNumbr As QuotePresenterErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As QuotePresenterErrors = -1 _
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
