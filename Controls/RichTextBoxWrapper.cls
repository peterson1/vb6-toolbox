VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RichTextBoxWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' http://computer-programming-forum.com/16-visual-basic/8432e721227bb14f.htm
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_LINESCROLL = &HB6


Const Lf$ = vbCrLf
Const Lf2$ = vbCrLf & vbCrLf


Private Enum RichTextBoxWrapperErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As RichTextBoxWrapperErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mFindString$, mSaveRequestInterval&, mLastSaveRequest!
Private mLastValue$ ' this doubles memory consumption

Private WithEvents mControl As RichTextBox
Attribute mControl.VB_VarHelpID = -1

Event KeyDown(KeyCode As Integer, Shift As Integer)
Event SaveRequest()




Private Sub mControl_Change()
    On Error GoTo ErrH
    
    '  don't bother if unchanged
    '
    If mControl.Text = mLastValue Then Exit Sub
    mLastValue = mControl.Text
    
    
    '  so we don't trigger Change too often
    '
    If Timer - mLastSaveRequest > mSaveRequestInterval Then
        mLastSaveRequest = Timer
        RaiseEvent SaveRequest
    End If
    
ErrH: Blame "mControl_Change"
End Sub


Private Sub mControl_LostFocus()
    On Error GoTo ErrH
    
    mLastSaveRequest = 0
    
    Call mControl_Change
    
ErrH: Blame "mControl_LostFocus"
End Sub


Private Sub mControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim origKey%
    On Error GoTo ErrH
    
    If Shift = 0 Then       ' no modifier keys pressed
        Select Case KeyCode
        
            Case vbKeyF3:   Find mFindString, False
            
            Case Else:      RaiseEvent KeyDown(KeyCode, Shift)
        End Select
        
        Exit Sub
    End If
    
    
    
    origKey = KeyCode
    
    If Shift = 2 Then 'Ctrl held down
        KeyCode = 0
        Select Case origKey
            
            Case vbKeyE: '  trap ctrl+E to keep from format center-aligned
            
            Case vbKeyF
                If LenB(Me.SelectedText) <> 0 Then mFindString = Me.SelectedText
                Find InputBox("Find this text:", "Search", mFindString), True
            
            Case vbKeyI: '  trap ctrl+I to keep from inserting Tabs
            
            
            Case vbKeyS:    RaiseEvent SaveRequest
            
            Case vbKeyT: '  trap ctrl+T to keep from ?
                            
            
            'Case vbKeyU:    MsgBox Me.SelectedText' testing lang
            
            
            '  trap ctrl+V to restrict to text pastes
            '
            Case vbKeyV:
                If Clipboard.GetFormat(vbCFText) Then _
                    Me.InsertAtCursor Clipboard.GetText
            
            
            '  if key is not notorious, let it be
            '
            Case Else:  KeyCode = origKey
            
        End Select
    End If
    
    
    RaiseEvent KeyDown(origKey, Shift)
    
    
    
ErrH: Blame "mControl_KeyDown"
End Sub



Public Property Get SelectedText() As String
    Dim startPos&, endPos&
    On Error GoTo ErrH
    With mControl
    
        startPos = .SelStart
        endPos = startPos + .SelLength
        
        SelectedText = Mid$(.Text, startPos + 1 _
                                 , endPos - startPos)
    End With
ErrH: Blame "[Get]SelectedText"
End Property


Public Property Let SearchString(lookForThisStr As String)
    mFindString = lookForThisStr
End Property

Public Sub Find(lookForThisStr As String _
              , startSearchFrmTop As Boolean)
    Dim startPos&
    On Error GoTo ErrH
    
    If LenB(lookForThisStr) = 0 Then Exit Sub
    
    Me.SearchString = lookForThisStr
    
    With mControl
        
        startPos = IIf(startSearchFrmTop, 1 _
                    , .SelStart + 1 + Len(mFindString))
        
        startPos = InStr(startPos, .Text, mFindString, vbTextCompare)
        
        If startPos = 0 Then
            If startSearchFrmTop Then
                MsgBox "Text does not contain:  """ & mFindString & """." _
                      , vbInformation, "Search string not found"
            Else
                If MsgBox("We reached the bottom of the text." & Lf2 _
                        & "Shall we start looking from the top?" _
                        , vbQuestion + vbOKCancel, "End of text reached") = vbOK _
                    Then Find mFindString, True
            End If
            Exit Sub
        End If
        
        .SelStart = startPos - 1
        .SelLength = Len(mFindString)
    End With
    
ErrH: Blame "Find"
End Sub




Public Sub InsertAtCursor(strPrefix As String _
                        , Optional strSuffix As String _
                        , Optional keepSelectd As Boolean = True _
                        )
    Dim startPos&, endPos&, origTxt$, scrollPos&
    On Error GoTo ErrH
    With mControl
        
        startPos = .SelStart
        endPos = startPos + .SelLength
        origTxt = .Text
        
        If LenB(Trim(origTxt)) = 0 Then
            .Text = strPrefix & strSuffix
            Exit Sub
        End If
        
        ' save scrollbar position
        scrollPos = SendMessage(mControl.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
        
        
        '  if no suffix....
        '
        If LenB(strSuffix) = 0 Then
            .Text = VBA.Left$(origTxt, startPos) _
                  & strPrefix _
                  & Mid$(origTxt, startPos + 1)
        
        
        '  if has suffix....
        '
        Else
            .Text = VBA.Left$(origTxt, startPos) _
                  & strPrefix _
                  & Mid$(origTxt, startPos + 1, endPos - startPos) _
                  & strSuffix _
                  & Mid$(origTxt, endPos + 1)
        End If
        
        ' restore scrollbar position
        Call SendMessage(mControl.hwnd, EM_LINESCROLL, 0, ByVal (scrollPos))
        
        If keepSelectd Then
            .SelStart = startPos + Len(strPrefix)
            .SelLength = endPos - startPos
        Else
            .SelStart = endPos + Len(strSuffix) + 2
        End If
        
        .SetFocus
        
    End With
ErrH: Blame "InsertAtCursor"
End Sub







Public Property Let PlainText(strPlainText As String)
    On Error GoTo ErrH
    
    '  because of this line, Property Let will NOT trigger change event.
    '
    mLastValue = strPlainText
    
    mControl.Text = strPlainText
    
ErrH: Blame "[Let]PlainText"
End Property

Public Property Get PlainText() As String
    On Error GoTo ErrH
    
    PlainText = mControl.Text
    
ErrH: Blame "[Get]PlainText"
End Property




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(richTextBoxCtrl As RichTextBox _
              , Optional saveRequestIntervalSecs As Long = 30 _
              )
    On Error GoTo ErrH
    
    mSaveRequestInterval = saveRequestIntervalSecs
    Set mControl = richTextBoxCtrl
    
    ' Redirect messages to our new handler in the module and
    ' save the pointer to the old handler
    pOldWindPoc = SetWindowLong(mControl.hwnd, GWL_WNDPROC, AddressOf WndProc)
    
    mLastSaveRequest = Timer
    
ErrH: Blame "Init"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Pass-through Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Let Top(topPostion As Long)
    mControl.Top = topPostion
End Property

Public Property Get Top() As Long
    Top = mControl.Top
End Property

Public Property Let Left(leftPostion As Long)
    mControl.Left = leftPostion
End Property

Public Property Get Left() As Long
    Left = mControl.Left
End Property

Public Property Get Height() As Long
    Height = mControl.Height
End Property

Public Property Let Height(ctrlHeight As Long)
    mControl.Height = ctrlHeight
End Property

Public Property Get Width() As Long
    Width = mControl.Width
End Property

Public Property Let Width(ctrlWidth As Long)
    mControl.Width = ctrlWidth
End Property

Public Property Get SelStart() As Long
    SelStart = mControl.SelStart
End Property

Public Property Let SelStart(startPos As Long)
    mControl.SelStart = startPos
End Property

Public Property Get SelLength() As Long
    SelLength = mControl.SelLength
End Property

Public Property Let SelLength(lengthOfSelection As Long)
    mControl.SelLength = lengthOfSelection
End Property





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Cleanup
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Call Me.Cleanup
    
ErrH: Blame "Class_Terminate"
End Sub

Public Sub Cleanup()
    On Error GoTo ErrH
    
    ' Restore the previously saved message handler
    If Not mControl Is Nothing Then _
        SetWindowLong mControl.hwnd, GWL_WNDPROC, pOldWindPoc
    
    Set mControl = Nothing
    
ErrH: Blame "Cleanup"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As RichTextBoxWrapperErrors = -1 _
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
                , Optional ByVal errorNumbr As RichTextBoxWrapperErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As RichTextBoxWrapperErrors = -1 _
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
