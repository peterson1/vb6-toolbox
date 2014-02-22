VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextEditorController"
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

Public Enum MarkupLanguages
    markup_None
    markup_HTML
    markup_Markdown
End Enum

Private Enum TextEditorControllerErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As TextEditorControllerErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private WithEvents mTextBox As VB.TextBox
Attribute mTextBox.VB_VarHelpID = -1

Private mLastUpdate!, mUpdateInterval& _
      , mMarkup As MarkupLanguages, mFindString$

Private mEndNotesMarker$(), mEndNotesText$(), mNextEndNoteIndex&

Event IntervalElapsed()
Event SaveRequested()


Private Sub mTextBox_Change()
    On Error GoTo ErrH
    
    If Timer - mLastUpdate > mUpdateInterval Then
        mLastUpdate = Timer
        RaiseEvent IntervalElapsed
    End If
    
    
ErrH: Blame "mTextBox_Change"
End Sub

Private Sub mTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    'Debug.Print Shift, KeyCode
    
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF3:       Find mFindString, False
        End Select
    
    ElseIf Shift = 2 Then 'Ctrl held down
        Select Case KeyCode
        
            Case vbKeyB:    ApplyBold
            Case vbKeyI:    ApplyItalic
            Case vbKeyE:    ApplyEndnote
            Case vbKeyA:    SelectAll
            Case vbKeyS:    RaiseEvent SaveRequested
            Case vbKeyF:    Find InputBox("Find this text:", "Search"), True
            
        End Select
    
    ElseIf Shift = 3 Then 'Ctrl+Shift held down
        Select Case KeyCode
        
        Case vbKeyE:    ApplyEndnoteFromBank
        Case vbKeyF:    FindEndNotes
        'Case vbKeyV:    PasteProcessed
        'Case vbKeyV:    Debug.Print "KeyDown"
        
        End Select
    End If
    
ErrH: Blame "mTextBox_KeyDown"
End Sub

Private Sub mTextBox_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    Select Case KeyAscii
        
        ' prevents ctrl+I from entering Tab char
        Case vbKeyTab:      KeyAscii = 0
        
        'Case vbKeyV:        Debug.Print "KeyPress Ctrl+V"
        
    End Select
ErrH: Blame "mTextBox_KeyPress"
End Sub

Private Sub mTextBox_LostFocus()
    On Error GoTo ErrH
    
    mLastUpdate = 0
    
    RaiseEvent IntervalElapsed
    
ErrH: Blame "mTextBox_LostFocus"
End Sub


'Private Sub PasteProcessed()
'    Dim clpbrd$
'    On Error GoTo ErrH
'
'    If Not Clipboard.GetFormat(vbCFText) Then Exit Sub
'    clpbrd = Clipboard.GetText
'
'    If LenB(clpbrd) = 0 Then Exit Sub
'
'    clpbrd = Replace(clpbrd, vbCrLf & vbTab, Lf2)
'
'    InsertAtCursor clpbrd
'
'ErrH: Blame "PasteProcessed"
'End Sub


Private Sub Find(lookForThisStr As String _
               , startSearchFrmTop As Boolean)
    Dim startPos&
    On Error GoTo ErrH
    
    If LenB(lookForThisStr) = 0 Then Exit Sub
    
    mFindString = lookForThisStr
    
    With mTextBox
        
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


Private Sub FindEndNotes()
    Dim clipbrd$, i&, ss$()
    On Error GoTo ErrH
    
    If Not Clipboard.GetFormat(vbCFText) Then Exit Sub
    clipbrd = Clipboard.GetText
    
    ' if not Tab-delimited, ignore
    If Not InText(clipbrd, vbTab) Then Exit Sub
    
    
    '  if multiple lines...
    If InText(clipbrd, vbCrLf) Then
        mEndNotesText = Split(clipbrd, vbCrLf)
        ReDim mEndNotesMarker(UBound(mEndNotesText))
        For i = 0 To UBound(mEndNotesText)
            
            ss = Split(mEndNotesText(i), vbTab)
            If UBound(ss) <> 1 Then
                MsgBox "Invalid tab-delimited endnote:" & Lf2 & mEndNotesText(i), vbExclamation
                Exit Sub
            End If
            
            mEndNotesMarker(i) = ss(0)
            mEndNotesText(i) = ss(1)
        Next i
    
    ' if single line
    Else
        ReDim mEndNotesMarker(0), mEndNotesText(0)
        
        ss = Split(clipbrd, vbTab)
        If UBound(ss) <> 1 Then
            MsgBox "Invalid tab-delimited endnote:" & Lf2 & mEndNotesText(i), vbExclamation
            Exit Sub
        End If
        
        mEndNotesMarker(0) = ss(0)
        mEndNotesText(0) = ss(1)
    End If
    
    mNextEndNoteIndex = 0
    Find mEndNotesMarker(mNextEndNoteIndex), True
    
ErrH: Blame "FindEndNotes"
End Sub



Private Sub ApplyBold()
    On Error GoTo ErrH
    Select Case mMarkup

        Case markup_HTML:       InsertAtCursor "<b>", "</b>"
        
        Case markup_Markdown:   InsertAtCursor "**", "**"

    End Select
ErrH: Blame "ApplyBold"
End Sub

Private Sub ApplyItalic()
    On Error GoTo ErrH
    Select Case mMarkup

        Case markup_HTML:       InsertAtCursor "<i>", "</i>"
        
        Case markup_Markdown:   InsertAtCursor "*", "*"

    End Select
ErrH: Blame "ApplyItalic"
End Sub


Private Sub InsertAtCursor(strPrefix As String _
                         , Optional strSuffix As String _
                         , Optional keepSelectd As Boolean = True _
                         )
    Dim startPos&, endPos&, origTxt$, scrollPos&
    On Error GoTo ErrH
    With mTextBox
        
        startPos = .SelStart
        endPos = startPos + .SelLength
        origTxt = .Text
        
        If LenB(Trim(origTxt)) = 0 Then
            .Text = strPrefix & strSuffix
            Exit Sub
        End If
        
        ' save scrollbar position
        scrollPos = SendMessage(mTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
        
        
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
        Call SendMessage(mTextBox.hwnd, EM_LINESCROLL, 0, ByVal (scrollPos))
        
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


Private Sub ApplyEndnote(Optional endNoteTextStr As String)
    Dim selectd$, startPos&, insertStr$
    On Error GoTo ErrH
    Select Case mMarkup

        'Case markup_HTML:       InsertAtCursor "<b>", "</b>"
        
        Case markup_Markdown:
'            InsertAtCursor "[^", "]" _
'                   & Lf2 & "[^" & Me.SelectedText & "]: " _
'                   , keepSelectd:=False
            
            ' surround selected text with brackets
            selectd = Me.SelectedText
            InsertAtCursor "[^", "]", keepSelectd:=True
            
            ' look for next line break
            With mTextBox
                startPos = InStr(.SelStart + Len(selectd), .Text, vbCrLf)
                If startPos = 0 Then startPos = .SelStart + Len(selectd)
                
                .SelStart = startPos + 2
                insertStr = "[^" & selectd & "]: "
                InsertAtCursor Lf & insertStr & Lf2, keepSelectd:=False
                
                .SelStart = startPos + 2 + Len(insertStr) + 2
            End With
            
    End Select
    
    If LenB(endNoteTextStr) <> 0 Then
        With mTextBox
        startPos = .SelStart
        
        InsertAtCursor endNoteTextStr
        
        .SelStart = startPos
        .SelLength = Len(endNoteTextStr)
        End With
    End If
    
ErrH: Blame "ApplyEndnote"
End Sub


Private Sub ApplyEndnoteFromBank()
    On Error GoTo ErrH
    
    Call ApplyEndnote(mEndNotesText(mNextEndNoteIndex))
    
    mNextEndNoteIndex = mNextEndNoteIndex + 1
    
    If mNextEndNoteIndex <= UBound(mEndNotesMarker) Then
        mFindString = mEndNotesMarker(mNextEndNoteIndex)
    Else
        mFindString = vbNullString
    End If
    
ErrH: Blame "ApplyEndnoteFromBank"
End Sub


Public Property Get SelectedText() As String
    Dim startPos&, endPos&
    On Error GoTo ErrH
    With mTextBox
    
        startPos = .SelStart
        endPos = startPos + .SelLength
        
        SelectedText = Mid$(.Text, startPos + 1 _
                                 , endPos - startPos)
    End With
ErrH: Blame "SelectedText"
End Property


Private Sub SelectAll()
    On Error GoTo ErrH
    With mTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
ErrH: Blame "SelectAll"
End Sub




'Private Function ScrollText&(textBoxCtrl As Control _
'                           , vLines As Long _
'                           )
'    #If Win32 Then
'        Dim scrolldLineCount&, savedWnd&, r&
'        Const EM_LINESCROLL = &HB6
'    #Else
'        Dim scrolldLineCount%, savedWnd%, r%
'        Const EM_LINESCROLL = &H406
'    #End If
'
'    ' Get the window handle of the control that currently has the
'    '  focus, Command1 or Command2.
'    savedWnd = Screen.ActiveControl.hwnd
'
'    ' Set the focus to the passed control (text control).
'    textBoxCtrl.SetFocus
'
'    ' Scroll the lines.
'    scrolldLineCount = SendMessage(textBoxCtrl.hwnd, EM_LINESCROLL, 0, vLines)
'
'    ' Restore the focus to the original control, Command1 or
'    '  Command2.
'
'    r = PutFocus(savedWnd)
'
'    ' Return the number of lines actually scrolled.
'    ScrollText& = scrolldLineCount
'
'End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(textBoxCtrl As VB.TextBox _
              , markupLanguageTyp As MarkupLanguages _
              , Optional minimumUpdateIntervalSecnds As Long = 30 _
              )
    On Error GoTo ErrH
    
    Set mTextBox = textBoxCtrl
    
    mMarkup = markupLanguageTyp
    
    mUpdateInterval = minimumUpdateIntervalSecnds
    
ErrH: Blame "Init"
End Sub









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo ErrH
    
    'Set mSomeObject = New Something
    
ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Set mTextBox = Nothing
    
    Erase mEndNotesMarker, mEndNotesText
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As TextEditorControllerErrors = -1 _
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
                , Optional ByVal errorNumbr As TextEditorControllerErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As TextEditorControllerErrors = -1 _
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

        'Err.Raise errorNumbr, .Source, errorDescrption
        MsgBox Err.Description, vbExclamation, .Source
        
    End With
End Sub
