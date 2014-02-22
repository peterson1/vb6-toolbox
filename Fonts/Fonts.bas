Attribute VB_Name = "Fonts"
Option Explicit
Const MODULE_NAME$ = "Fonts"

Private Enum FontsErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As FontsErrors    '
    Description As String
End Type
Private mError As ErrorHolder



Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Private Sub Form_Load()
'    Dim res As Long
'    ' add the font
'    res = AddFontResource("C:\Fonts\Nordic__.ttf")
'    If res > 0 Then
'        ' alert all windows that a font was added
'        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
'        MsgBox res & " fonts were added!"
'    End If
'End Sub



' http://www.answers.com/topic/addfontresource#ixzz28OZFIhAo
Public Sub Add(pathOfFontFile As String)
    On Error GoTo ErrH
    
    Const HWND_BROADCAST = &HFFFF&
    Const WM_FONTCHANGE = &H1D

    If AddFontResource(pathOfFontFile) > 0 Then
        ' alert all windows that a font was added
        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
    End If
    
ErrH: Blame "Add"
End Sub

' http://www.answers.com/topic/removefontresource#ixzz28ObHwbqn
Public Sub Remove(pathOfFontFile As String)
    On Error GoTo ErrH
    
    Call RemoveFontResource(pathOfFontFile)
    
ErrH: Blame "Remove"
End Sub








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As FontsErrors = -1 _
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

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As FontsErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As FontsErrors = -1 _
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
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub

