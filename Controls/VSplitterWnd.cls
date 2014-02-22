VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVSplitterWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Brad Martinez,  http://www.mvps.org/ccrp/
'
' =================================================
' A vertical splitter bar class...
'
' 04-13-99:
'   - Added Top property
'   - Added HeightOffset property, allows the splitter's dynamic "bottom
'     position" to be set.
'   - Added proportional splitter positioning, where the spiltter maintains
'     it's relative position in the parent widow when it is resized.
'   - Changed the two module level Control variables Longs, now issues
'     MoveWindow instead of the Move method on these varables; this
'     simplifies and makes things more efficient, but now VB controls without
'     no window handle can no longer be assigned to these variables
'     (Image, Label, Shape); changed the name of this class accordingly.

' 06-07-99:
'   - Added Percent property, allowing the splitter position to be set
'     as a percentage of the parent window's ScaleWidth.
'   - Renamed the ParentResize proc to AdjustSplitterPosition, more
'     accurately reflecting what the proc does.
'
' =================================================

Private Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const WS_BORDER = &H800000
Private Const WS_EX_CLIENTEDGE = &H200

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As CBoolean) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long

Private Const HWND_DESKTOP = 0

Private Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Private Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As GetSystemMetrics_codes) As Long
Private Enum GetSystemMetrics_codes
  SM_CXBORDER = 5
'  SM_CYBORDER = 6
  SM_CXFRAME = 32
'  SM_CYFRAME = 33
  SM_CXMINTRACK = 34
'  SM_CYMINTRACK = 35
  SM_CXSIZEFRAME = SM_CXFRAME               ' win40 name change
'  SM_CYSIZEFRAME = SM_CYFRAME               ' win40 name change
End Enum

' ==========================================================
' Required in the client for proper splitter functionality:
'
' The Parent (Form, UserControl), with the following default property settings
' changed to:
'    KeyPreview = True
'    ScaleMode = vbPixels
'
' Two child control windows in the parent, both having an hWnd property,
' and similar Appearance and BorderStyle settings.

' A Frame control, with the following default property settings changed to:
'    Appearance = 0   ' flat
'    BorderStyle = vbBSNone
'    ClipControls = False (optional)
'
' ==========================================================

' The minimum allowable window width of the child controls, in pixels.
Private m_cxMinChild As Long

' The splitter bar's horizontal and vertical offsets, in pixels.
Private m_cxOffset As Long   ' WS_BORDER
Private m_cyOffset As Long   ' distance from bottom of parent form

' Splitter bar's width, in pixels
Private m_cxSplit As Long

' Splitter bar's current x and y positions, relative to the
' parent form's client coords, in pixels
Private m_xSplit As Long
Private m_ySplit As Long

' The parent's current ScaleWidth and ScaleHeight, in pixels
Private m_cxClient As Long
Private m_cyClient As Long

' for proportional resizing, the ratio is used only when the parent form is resized
Private m_cxClientStart As Long
Private m_sngSplitRatio As Single

' for track splitting. m_xSplitStart is used to restore the splitter's initial position
' if a track split is cancelled.
Private m_fTrackSplit As Boolean
Private m_xSplitStart As Long

' The desktop's hDC. Used by PatBlt to draw the splitter bar, and is
' used as a flag to indicate when we're currently splitting.
Private m_hdcDesktop As Long

Private m_hwndParent As Long
Attribute m_hwndParent.VB_VarHelpID = -1
Private m_hwndCtrl1 As Long   ' the left-hand control
Private m_hwndCtrl2 As Long   ' the right-hand control
Private m_fraSplit As Frame
'

Private Sub Class_Terminate()
  
  ' If we have it, release the desktop DC.
  ' The variable is zeroed when released elsewhere.
  If m_hdcDesktop Then
    Call ReleaseDC(HWND_DESKTOP, m_hdcDesktop)
  End If

  Set m_fraSplit = Nothing

End Sub

' ================================================================
' Public members

Public Sub SetControls(hwndParent As Long, hwndCtrl1 As Long, hwndCtrl2 As Long, fra As Frame)
    
  m_hwndParent = hwndParent
  m_hwndCtrl1 = hwndCtrl1
  m_hwndCtrl2 = hwndCtrl2
  Set m_fraSplit = fra

  ' Setup the splitter frame control. It serves 2 purposes here:
  '   - Shows the splitter cursor when the mouse is over it
  '   - Allows the splitter cursor to appear in a 2 pixel wider area than
  '     the distance betwen the two controls.
  With m_fraSplit
'    .Appearance = 0   ' flat
    .BackColor = vbActiveBorder
'    .BorderStyle = vbBSNone
    .MousePointer = vbSizeWE
    .ZOrder   ' put the frame on top
  End With

  ' needs m_hwndParent
  Call FillSplitterMetrics
    
End Sub

Public Property Get Percent() As Byte
  Dim rc As RECT
    
  If m_hwndParent Then
    Call GetClientRect(m_hwndParent, rc)
    Percent = (m_xSplit / (rc.Right - rc.Left)) * 100
  End If
  
End Property

Public Property Let Percent(b As Byte)
  Dim rc As RECT
    
  If m_hwndParent Then
    ' Clear the proportional split ratio (will be set in AdjustSplitterPosition)
    m_sngSplitRatio = 0
    Call GetClientRect(m_hwndParent, rc)
    m_xSplit = (rc.Right - rc.Left) * (b * 0.01)
    Call AdjustSplitterPosition(True)
  End If

End Property

Public Property Get Left() As Long
  Left = m_xSplit
End Property

Public Property Let Left(x As Long)
  If (x >= 0) Then
    m_xSplit = x
    ' Clear the proportional split ratio (will be set in AdjustSplitterPosition)
    m_sngSplitRatio = 0
    Call AdjustSplitterPosition(True)
  End If
End Property

Public Property Get Width() As Byte
  Width = m_cxSplit
End Property

Public Property Let Width(cx As Byte)
  If cx Then
    m_cxSplit = cx
    Call AdjustSplitterPosition(True)
  End If
End Property

Public Property Get Top() As Long
  Top = m_ySplit
End Property

Public Property Let Top(y As Long)
  If (y >= 0) Then
    m_ySplit = y
    Call AdjustSplitterPosition(True)
  End If
End Property

Public Property Get HeightOffset() As Long
  HeightOffset = m_cyOffset
End Property

Public Property Let HeightOffset(cy As Long)
  If (cy >= 0) Then
    m_cyOffset = cy
    Call AdjustSplitterPosition(True)
  End If
End Property

' Adjusts the splitter bar position. Was ParentResize

Public Sub AdjustSplitterPosition(fUpdateCtrls As Boolean)
  Dim rc As RECT
'  On Error GoTo Out
  
  If m_hwndParent Then
    
    ' Get the parent window's current client rect
    Call GetClientRect(m_hwndParent, rc)
  
    ' Establish the ratio of the splitter's current x position relative to the parent's
    ' previous ScaleWidth. On split start m_sngSplitRatio is cleared, and
    ' m_cxClientStart is reinitialized to the current ScaleWidth.
    ' The "Val" function is used because a Single data type can contain the
    ' overflow value "-1.#IND", and will err anytime the variable's value is
    ' referenced directly when it holds this value.
    If (val(m_sngSplitRatio) <= 0) Then
      If (m_xSplit = 0) Then m_xSplit = m_fraSplit.Left
      If (m_cxClientStart = 0) Then m_cxClientStart = rc.Right - rc.Left   ' frmParent.ScaleWidth
      m_sngSplitRatio = m_xSplit / m_cxClientStart
    End If
      
    ' Get the parent's client area in the mod level variables. Set
    ' the client height so that it's relative to the height offset here.
    m_cxClient = rc.Right - rc.Left   ' frmParent.ScaleWidth
    m_cyClient = (rc.Bottom - rc.Top) - m_cyOffset   ' frmParent.ScaleHeight - m_cyOffset
    
    ' Move the splitter only if the relative parent client height is
    ' equal to or greater than the splitter's height (the frame ctrl is
    ' the splitter bar, and we'll err if we make its height negative)
    If (m_cyClient >= m_ySplit) Then
      ' Calculate the splitter bar's left coordinate using the ratio
      m_xSplit = Min(Max((m_cxMinChild), m_cxClient * m_sngSplitRatio), m_cxClient - (m_cxSplit + m_cxMinChild))
      Call MoveControls(fUpdateCtrls)
    End If
  
  End If   ' m_hwndParent
  
Out:
End Sub

Public Property Get Splitting() As Boolean
  Splitting = m_hdcDesktop
End Property

Public Property Get TrackSplit() As Boolean
  TrackSplit = m_fTrackSplit
End Property

Public Property Let TrackSplit(fTrackSplit As Boolean)
  m_fTrackSplit = fTrackSplit
End Property

' Initiates splitter bar dragging.
' the x and y params are the frame's internal clent coords, in TWIPS!!!

Public Sub BeginSplit(x As Single, y As Single)
  Dim rc As RECT
  
  ' Get the desktop's DC (also used as a flag indicating we're currently splitting)
  m_hdcDesktop = GetDC(HWND_DESKTOP)
  If m_hdcDesktop Then
    
    ' Store the parent's current scale width and clear the
    ' proportional split ratio (both are set in AdjustSplitterPosition)
    Call GetClientRect(m_hwndParent, rc)
    m_cxClientStart = rc.Right - rc.Left
    m_sngSplitRatio = 0
        
    ' Calculate the splitter bar's initial position. The splitter's left coordinate
    ' will never be less than the first control's minimum control width (-1 pixel to
    ' cover the first control's right border), and never be more than the second control's
    ' left coordinate (minus the splitter's width + 1 pixel to cover the second control's
    ' left border). There's a 1 pixel slop when the form is at it's minimum width...
    m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left + (x \ Screen.TwipsPerPixelX)), _
                           m_cxClient - (m_cxSplit + m_cxMinChild))
    
    If m_fTrackSplit Then
      m_xSplitStart = m_xSplit
    Else
      ' Draw the splitter bar on top of everything.
      Call DrawSplitter(m_xSplit, m_ySplit, m_cxSplit, m_cyClient - m_ySplit)
    End If
    
  End If   ' m_hdcDesktop
  
End Sub

' Moves the splitter bar
' the x and y params are the frame's internal clent coords, in TWIPS!!!

Public Sub MoveSplitter(x As Single, y As Single)
  Dim cx As Long
  
  If m_hdcDesktop Then
    
    If (m_fTrackSplit = False) Then
      ' Erase the previous splitter bar.
      Call DrawSplitter(m_xSplit, m_ySplit, m_cxSplit, m_cyClient - m_ySplit)
    End If
    
    ' Calculate the splitter bar's new position.
    m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left + (x \ Screen.TwipsPerPixelX)), _
                           m_cxClient - (m_cxSplit + m_cxMinChild))
  
    If m_fTrackSplit Then
      Call MoveControls(True)
    Else
      ' Draw the splitter bar's new position.
      Call DrawSplitter(m_xSplit, m_ySplit, m_cxSplit, m_cyClient - m_ySplit)
    End If
    
  End If
End Sub

' Finishes splitter bar moving and resizes the form's child windows.
' the x and y params are the frame's internal clent coords, in TWIPS!!!

Public Sub EndSplit(x As Single, y As Single, fUpdateCtrls As Boolean)
  Dim wSplitRight As Long
  
  If m_hdcDesktop Then
    
    If (m_fTrackSplit = False) Then
      ' Erase the previous bar.
      Call DrawSplitter(m_xSplit, m_ySplit, m_cxSplit, m_cyClient - m_ySplit)
    End If
    
    ' Release and clear the desktop's hDC variable.
    Call ReleaseDC(HWND_DESKTOP, m_hdcDesktop)
    m_hdcDesktop = 0
    
    ' Calculate the splitter's final position.
    m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left + (x \ Screen.TwipsPerPixelX)), _
                           m_cxClient - (m_cxSplit + m_cxMinChild))
    
    ' Move the controls to their new positions.
    Call MoveControls(fUpdateCtrls)
    
  End If
End Sub

' Cancels any current splitting operation.

Public Sub CancelSplit()

  If m_hdcDesktop Then

    ' Remove the splitter pointer.
    Call ReleaseCapture
    
    ' Restore the splitter's original position.
    If m_fTrackSplit Then
      m_xSplit = m_xSplitStart
      Call MoveControls(True)
    Else
      ' Erase the previous splitter bar.
      Call DrawSplitter(m_xSplit, m_ySplit, m_cxSplit, m_cyClient - m_ySplit)
    End If
    
    ' Release and clear the desktop's hDC variable.
    Call ReleaseDC(HWND_DESKTOP, m_hdcDesktop)
    m_hdcDesktop = 0
  
  End If

End Sub

' ================================================================
' Private calls

' Sets the splitter's width to match the window sizing border width (which will be a
' width of 2 WS_BORDER's greater than the distance between the controls). Also
' sets the minimum width for the controls (m_cxMinChild)

' A top level window's default minimum width (SM_CXMINTRACK) minus the widths
' of it's left and right sizing borders (SM_CXSIZEFRAME), leaves a constant minimum
' client width of 104 pixels.

Private Sub FillSplitterMetrics()
  Dim cxSizeFrame As Long
  Dim cxBorder As Long
  Dim cxMinTrack As Long
  Dim dwStyle As Long
  
  If m_hwndParent Then
    
    ' Get the top level window sizing border width (default is 4 pixels)
    cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME)
    
    ' Get the window border width (WS_BORDER, default is 1 pixel)
    cxBorder = GetSystemMetrics(SM_CXBORDER)
  
    ' Set the splitter bar's width (the top level window sizing border
    ' width minus the window border width, default is 3 pixels)
    If (m_cxSplit = 0) Then
      m_cxSplit = cxSizeFrame - cxBorder
    End If
    
    ' Set the control offset. The splitter will cover a border width
    ' of the each control only when they have both of their
    ' WS_BORDER and WS_EX_CLIENTEDGE style bits set.
    ' (Appearance and BorderStyle properties respectively).
    ' We'll just check the first control, and assume the second
    ' is set the same...
    If ((GetWindowLong(m_hwndCtrl1, GWL_STYLE) And WS_BORDER) = WS_BORDER) And _
       ((GetWindowLong(m_hwndCtrl1, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE) Then
  '  If (Ctrl1.Appearance = 3D) And (Ctrl1.BorderStyle = FixedSingle) Then
      m_cxOffset = cxBorder
    End If
    
   ' Get the top level window minimum tracking width.
    cxMinTrack = GetSystemMetrics(SM_CXMINTRACK)
    
    ' Calculate the control's minimum horizontal tracking width.
    ' (usually around 50 pixels)
    m_cxMinChild = (cxMinTrack - ((cxSizeFrame * 2) + m_cxSplit)) \ 2
  
  End If   ' m_hwndParent
  
End Sub

' only called when m_fTrackSplit = False

Private Sub DrawSplitter(x As Long, y As Long, cx As Long, cy As Long)
  Dim pt As POINTAPI
  pt.x = x
  pt.y = y
  Call ClientToScreen(m_hwndParent, pt)
  Call PatBlt(m_hdcDesktop, pt.x, pt.y, cx, cy, vbDstInvert)
End Sub

Private Sub MoveControls(fUpdateCtrls As Boolean)
  Static xSplitPrev As Long
  Dim wSplitRight As Long
  On Error GoTo Out

  ' ======================================================
  ' Resize the first control.
  
  ' If the splitter's new position is less than it's previous position...
  If m_xSplit < xSplitPrev Then
    ' Resize m_hwndCtrl1 first and repaint it if specified.
    Call MoveWindow(m_hwndCtrl1, 0, m_ySplit, _
                                 m_xSplit + m_cxOffset, m_cyClient - m_ySplit, CTrue)
    If fUpdateCtrls Then Call UpdateWindow(m_hwndCtrl1)
  Else
    ' Resize m_hwndCtrl2 first and repaint it if specified.
    wSplitRight = m_xSplit + m_cxSplit
    Call MoveWindow(m_hwndCtrl2, wSplitRight - m_cxOffset, m_ySplit, _
                                  (m_cxClient - wSplitRight) + m_cxOffset, m_cyClient - m_ySplit, CTrue)
    If fUpdateCtrls Then Call UpdateWindow(m_hwndCtrl2)
  End If
    
  ' ======================================================
  ' Move the splitter bar (frame).
  
  ' Make the frame invisible before the move, or the frame's
  ' shadow may be left on the controls when splitting.
  m_fraSplit.Visible = False
  
  ' Center the frame between the two controls making it the full
  ' height of the parent. The frame's Move method is really slow...
'  m_fraSplit.Move m_xSplit, 0, m_cxSplit, m_cyClient
  Call MoveWindow(m_fraSplit.hWnd, m_xSplit, m_ySplit, m_cxSplit, m_cyClient - m_ySplit, CTrue)
  
  ' Make the frame visible again and repaint it if specified.
  m_fraSplit.Visible = True
  If fUpdateCtrls Then Call UpdateWindow(m_fraSplit.hWnd)

  ' ======================================================
  ' Resize the second control.
  
  If m_xSplit < xSplitPrev Then
    ' Resize m_hwndCtrl2 and repaint it if specified.
    wSplitRight = m_xSplit + m_cxSplit
    Call MoveWindow(m_hwndCtrl2, wSplitRight - m_cxOffset, m_ySplit, _
                                 (m_cxClient - wSplitRight) + m_cxOffset, m_cyClient - m_ySplit, CTrue)
    If fUpdateCtrls Then Call UpdateWindow(m_hwndCtrl2)
  Else
    ' Resize m_hwndCtrl1 and repaint it if specified.
    Call MoveWindow(m_hwndCtrl1, 0, m_ySplit, _
                                  m_xSplit + m_cxOffset, m_cyClient - m_ySplit, CTrue)
    If fUpdateCtrls Then Call UpdateWindow(m_hwndCtrl1)
  End If
  
Out:
  ' Save the splitter's current position.
  xSplitPrev = m_xSplit
  
End Sub

' Returns the smaller of the two specified values.

Private Function Min(param1 As Long, param2 As Long) As Long
  If param1 < param2 Then Min = param1 Else Min = param2
End Function

' Returns the larger of the two specified values.

Private Function Max(param1 As Long, param2 As Long) As Long
  If param1 > param2 Then Max = param1 Else Max = param2
End Function
