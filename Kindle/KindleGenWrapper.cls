VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KindleGenWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum KindleGenCompressionModes
    c0_NoCompression = 0
    c1_StandardDOC = 1
    c2_KindleHuffdic = 2
End Enum

Private Enum KindleGenWrapperErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    BuildError
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As KindleGenWrapperErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mKindleGenExe$, mPreviewerExe$


Public Sub MakeAndView(pathOfKindleOpf As String _
                     , Optional compressionMethd As KindleGenCompressionModes = c0_NoCompression _
                     , Optional nameOfmobiFile As String = "draft.mobi" _
                     )
    Dim outputF$
    On Error GoTo ErrH
    
    RunCommand "taskkill /IM javaw.exe /T", , , vbNormalFocus
    
    outputF = Me.Make(pathOfKindleOpf, compressionMethd, nameOfmobiFile)
    
    Shell """" & mPreviewerExe & """ """ & outputF & """", vbNormalFocus
    
ErrH: Blame "MakeAndView"
End Sub

Public Function Make(pathOfKindleOpf As String _
                   , Optional compressionMethd As KindleGenCompressionModes = c2_KindleHuffdic _
                   , Optional nameOfmobiFile As String = "draft.mobi" _
                   ) As String
    Dim cmd$, extCode&, outputF$, cmdOut$
    On Error GoTo Cleanup
    
    
    '  delete target file as needed
    '
    outputF = F_.Parent(pathOfKindleOpf) & nameOfmobiFile
    Call F_.Delete(outputF)
    
    
    '  compose command with args
    '
    cmd = """" & mKindleGenExe & """ " _
        & """" & pathOfKindleOpf & """ " _
        & "-c" & compressionMethd & " " _
        & "-o " & nameOfmobiFile
    
    
    '  run the command
    '
    extCode = F_.ShellWait(cmd, , , vbNormalFocus)

    
    If extCode = 0 Then
        ErrorIf Not F_.Found(outputF) _
             , "Failed to create """ & nameOfmobiFile & """, but exit code was [0]."
    
    ElseIf extCode <> 0 Then
        cmdOut = GetCommandOutput(cmd)
        Err.Raise BuildError, , GetError(cmdOut, extCode)
    End If
    
    
    '  return path of .mobi file
    '
    Make = outputF
    
    
Cleanup:    SaveError
            'F_.Delete redirectF
            'Erase someArray
  LoadError "Make" ', "details of error"
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(kindlGenExePath As String _
              , Optional kindlePreviewrExePath As String _
              )
    On Error GoTo ErrH
    
    
    ErrorIf Not Found(kindlGenExePath) _
         , "KindleGen.exe not found in:" & vbCrLf & kindlGenExePath
    mKindleGenExe = kindlGenExePath
    
    If LenB(kindlePreviewrExePath) <> 0 Then
        ErrorIf Not Found(kindlePreviewrExePath) _
             , "KindlePreviewer .exe not found in:" & vbCrLf & kindlePreviewrExePath
        mPreviewerExe = kindlePreviewrExePath
    End If
    
ErrH: Blame "Init"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function GetError(CmdOutput As String _
                        , extCode As Long _
                        ) As String
    Dim markr$, errStart&, errEnd&
    On Error GoTo ErrH
    
    Select Case extCode
        
        Case 1 'warnings
            markr = "Warning("
            errStart = InStr(CmdOutput, markr)
            errEnd = InStr(errStart, CmdOutput, vbCrLf)
            GetError = Mid$(CmdOutput, errStart + Len(markr), errEnd - errStart - Len(markr) - 1)
        
        Case 2 'error
            markr = "Error("
            errStart = InStrRev(CmdOutput, markr)
            GetError = Mid$(CmdOutput, errStart + Len(markr))
        
    End Select
    
ErrH: Blame "GetError"
End Function









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
    
    'Set mSomeObject = Nothing
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As KindleGenWrapperErrors = -1 _
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
                , Optional ByVal errorNumbr As KindleGenWrapperErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As KindleGenWrapperErrors = -1 _
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

