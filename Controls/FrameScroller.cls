VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum FrameScrollerErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As FrameScrollerErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mResizedHere As Boolean _
      , mOldBoundaryWidth&

Private WithEvents mTarget As VB.PictureBox _
      , WithEvents mBoundary As VB.PictureBox _
      , WithEvents mScrollBar As VB.VScrollBar
Attribute mTarget.VB_VarHelpID = -1
Attribute mBoundary.VB_VarHelpID = -1
Attribute mScrollBar.VB_VarHelpID = -1

Event ScrollBarToggled()


Public Sub Refresh()
    Dim prevVisbl As Boolean
    On Error GoTo ErrH
    
    If Not mTarget.Visible Then Exit Sub
    
    If mResizedHere Then Exit Sub   ' ignore if resize came from within
    mResizedHere = True             '
    
    With mScrollBar
        prevVisbl = .Visible
        .Visible = mTarget.Height > mBoundary.Height
        
        If .Visible Then
            
            If mOldBoundaryWidth <> mBoundary.Width Then    ' so we don't shrink
                mBoundary.Width = mBoundary.Width - .Width  '  if only target resized
                mOldBoundaryWidth = mBoundary.Width
            End If
            
            .Max = mTarget.ScaleHeight - mBoundary.ScaleHeight
            .LargeChange = mBoundary.ScaleHeight
            .SmallChange = mBoundary.ScaleHeight
            
            .Left = mBoundary.Left + mBoundary.Width
            .Top = mBoundary.Top
            .Height = mBoundary.Height
        End If
    
        If .Visible <> prevVisbl _
            Then RaiseEvent ScrollBarToggled
        
    End With
    
    mResizedHere = False
    
ErrH: Blame "Refresh"
End Sub


Private Sub mTarget_Resize()
    On Error GoTo ErrH
    
    Call Me.Refresh
    
ErrH: Blame "mTarget_Resize"
End Sub


Private Sub mBoundary_Resize()
    On Error GoTo ErrH
    
    Call Me.Refresh
    
ErrH: Blame "mBoundary_Resize"
End Sub

Private Sub mScrollBar_Change()
    Call ScrollVertical(mScrollBar.Value)
End Sub

Private Sub mScrollBar_Scroll()
    Call ScrollVertical(mScrollBar.Value)
End Sub

Private Sub mScrollBar_GotFocus()
    mBoundary.SetFocus
End Sub

Private Sub ScrollVertical(scrolBarValue As Long)
    On Error GoTo ErrH
    
    mTarget.Top = -scrolBarValue - 10
    
ErrH: Blame "ScrollVertical"
End Sub




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(picBoxToBeScrolled As VB.PictureBox _
              , boundingPicBox As VB.PictureBox _
              , vertcalScrollBar As VB.VScrollBar _
              )
    On Error GoTo ErrH
    
    Set mTarget = picBoxToBeScrolled
    Set mBoundary = boundingPicBox
    Set mScrollBar = vertcalScrollBar
    
    mTarget.Top = -23
    
    Call mBoundary_Resize
    
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
    
    Set mTarget = Nothing
    Set mBoundary = Nothing
    Set mScrollBar = Nothing
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As FrameScrollerErrors = -1 _
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
                , Optional ByVal errorNumbr As FrameScrollerErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As FrameScrollerErrors = -1 _
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
