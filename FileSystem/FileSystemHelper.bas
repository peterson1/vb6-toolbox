Attribute VB_Name = "F_"
Option Explicit
Const MODULE_NAME$ = "F_"

Const Lf$ = vbCrLf
Const Lf2$ = vbCrLf & vbCrLf

Private Enum F_Errors    ' you may make this Public for tests
    ErrorBase = 3000                ' you may adjust this minimum
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As F_Errors    '
    Description As String
End Type
Private mError As ErrorHolder

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long





Public Function BrowseForFolder(ownerFormHwnd As Long _
                              , Optional strPrompt As String = "Select which folder to use." _
                              , Optional rootFoldr As String _
                              , Optional warnIfHasContents As Boolean = False _
                              ) As String
    Dim foldr As Shell32.Folder2, contnts&
    On Error GoTo Cleanup
    
    Const foldrOnly& = &H1
    
    With New Shell32.Shell
        Set foldr = .BrowseForFolder(ownerFormHwnd, strPrompt _
                                   , foldrOnly, rootFoldr)
        
        If Not foldr Is Nothing Then _
            BrowseForFolder = foldr.Self.Path & "\"
    End With
    
    If warnIfHasContents Then
        Call List(BrowseForFolder, rsultCount:=contnts)
        If contnts <> 0 Then
            If MsgBox("Selected folder is not empty." & Lf _
                    & "Existing files may be overwritten." & Lf2 _
                    & "Are you sure you want to use this folder?" & Lf & BrowseForFolder _
                    , vbInformation + vbYesNo) <> vbYes Then BrowseForFolder = vbNullString
        End If
    End If
    
Cleanup:    SaveError
            Set foldr = Nothing
            'Erase someArray
  LoadError "BrowseForFolder" ', "details of error"
End Function






Public Function Found(fileOrFoldrPath As String _
                    ) As Boolean
    On Error GoTo Cleanup
    
    If FileFound(fileOrFoldrPath) Then
        Found = True
        
    ElseIf IsFolder(fileOrFoldrPath) Then
        Found = True
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Found" ', "details of error"
End Function


Public Sub Copy(sourceFileOrFoldr As String _
              , ByVal targetFileOrFoldr As String _
              )
    On Error GoTo Cleanup
    
    
    '  make sure source exists
    '
    ErrorIf Not Found(sourceFileOrFoldr) _
         , "File/folder to copy does not exist."
    
    
    
    
    '  if source if a folder,
    '   - assume target is a folder
    '   - recreate whole folder in target
    '
    If IsFolder(sourceFileOrFoldr) Then
        
        Call CopyFolderToFolder(sourceFileOrFoldr _
                              , targetFileOrFoldr)
    
    
    
    '  if source is a file...
    '
    Else
        '  if target already exists as a folder,
        '   - use filename of original file
        '   - copy file to existing folder
        '
        If IsFolder(targetFileOrFoldr) Then
            
            targetFileOrFoldr = Folderize(targetFileOrFoldr) _
                              & FileName(sourceFileOrFoldr)
            
            Call VBA.FileCopy(sourceFileOrFoldr _
                            , targetFileOrFoldr)
        
        
        
        '  if target already exists as a file,
        '   - delete existing file before copying
        '
        ElseIf FileFound(targetFileOrFoldr) Then

            Call Delete(targetFileOrFoldr)

            Call VBA.FileCopy(sourceFileOrFoldr _
                            , targetFileOrFoldr)



        '  if target doesn't exist,
        '   and it LOOKS LIKE a folder (ends in "\"),
        '    - create target folder first
        '    - use filename of original file
        '    - copy file to newly created folder
        '
        ElseIf Right$(RTrim$(targetFileOrFoldr), 1) = "\" Then

            Call CreateFolder(targetFileOrFoldr)

            targetFileOrFoldr = Folderize(targetFileOrFoldr) _
                              & FileName(sourceFileOrFoldr)

            Call VBA.FileCopy(sourceFileOrFoldr _
                            , targetFileOrFoldr)



        '  if target doesn't exist,
        '   and it doesn't look like a folder,
        '    - treat target as a non-existent file
        '    - create parent folders of non-existent target file
        '    - copy file to newly created folder
        '
        Else
            Call CreateFolder(Parent(targetFileOrFoldr))

            Call VBA.FileCopy(sourceFileOrFoldr _
                            , targetFileOrFoldr)
        End If
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Copy", "[source]: " & vbCrLf & sourceFileOrFoldr _
         & vbCrLf & "[target]: " & vbCrLf & targetFileOrFoldr
End Sub

Private Sub CopyFolderToFolder(sourceFoldr$ _
                             , targetFoldr$)
    On Error GoTo Cleanup
    
    Err.Raise 9999, , "Unfinished method"
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "CopyFolderToFolder" ', "details of error"
End Sub


Public Sub Move(srcFilePath As String _
              , destFilePath As String _
              )
    On Error GoTo ErrH
    
    ErrorIf Not F_.Found(srcFilePath) _
         , "Source file not found:" & Lf & srcFilePath
    
    Call F_.Copy(srcFilePath, destFilePath)
    
    Call F_.Delete(srcFilePath)
    
ErrH: Blame "Move"
End Sub


Public Function FileName(ByVal pathOfFile As String _
                       , Optional includeExtnsion As Boolean = True _
                       ) As String
    Dim pos&
    On Error GoTo Cleanup
    
    pathOfFile = RTrim$(pathOfFile)
    
    If Right$(pathOfFile, 1) = "\" Then Exit Function
    
    pos = InStrRev(pathOfFile, "\")
    
    FileName = Right$(pathOfFile, Len(pathOfFile) - pos)
    
    
    If Not includeExtnsion Then
        pos = InStr(FileName, ".")
        
        If pos <> 0 Then _
            FileName = Left$(FileName, pos - 1)
    End If
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "FileName" ', "details of error"
End Function



Public Function Parent(ByVal fileOrFoldrPath As String _
                     ) As String
    Dim pos&
    On Error GoTo Cleanup
    
    fileOrFoldrPath = RTrim$(fileOrFoldrPath)
    
    
    '  if ends in "\",
    '   ...
    '
    If Right$(fileOrFoldrPath, 1) = "\" Then
        pos = InStrRev(fileOrFoldrPath, "\", Len(fileOrFoldrPath) - 1)
    
    '  otherwise:
    Else
        pos = InStrRev(fileOrFoldrPath, "\")
    End If
    
    Parent = Left$(fileOrFoldrPath, pos)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Parent" ', "details of error"
End Function


Public Sub CreateFolder(foldrPath As String)
    Dim ss$(), i&, iMin&, iMax&, foldr$
    On Error GoTo Cleanup
    
    If IsFolder(foldrPath) Then Exit Sub
    
    ss = Split(foldrPath, "\")
    
    '  if network folder,
    '   - start with 4th token
    '   - otherwise, start with 2nd
    '
    iMin = IIf(IsRemote(foldrPath), 3, 1)
    iMax = UBound(ss)
    
    For i = iMin To iMax
        ss = Split(foldrPath, "\")
        ReDim Preserve ss(i)
        foldr = Join(ss, "\")
        
        If Not IsFolder(foldr) _
            Then Call MkDir(foldr)
    Next i
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "CreateFolder", "Can't create folder:" & vbCrLf & foldr
End Sub


Public Function CreateFile(pathOfFile$ _
                         , Optional fileContnts$ _
                         ) As String
    Dim fn&
    On Error GoTo Cleanup
    
    '  create containing folder(s)
    '
    Call CreateFolder(Parent(pathOfFile))
    
    
    fn = FreeFile
    Open pathOfFile For Output As #fn

    Print #fn, fileContnts;

    Close #fn
    
    '  return path of file
    CreateFile = pathOfFile
    
Cleanup:    SaveError
            'Set f = Nothing
            'Erase someArray
  LoadError "CreateFile" ', "details of error"
End Function



Public Sub Delete(fileOrFoldrPath As String)
    'Dim extCode&, msg$
    On Error GoTo Cleanup
    
    ErrorIf LenB(fileOrFoldrPath) = 0 _
         , "File/folder parameter cannot be blank."
    
    If IsFolder(fileOrFoldrPath) Then
        'Call RmDir(fileOrFoldrPath) <--- this won't work if has contents
        
'        extCode = RunCommand("RmDir /S /Q " _
'                     & """" & fileOrFoldrPath & """")
'
'        If extCode <> 0 Then
'            Select Case extCode
'                Case 32:    msg = "The process cannot access the file because it is being used by another process."
'                Case Else:  msg = "Unexpected RmDir error."
'            End Select
'            msg = "Unable to delete folder using RmDir." & vbCrLf & msg
'            Err.Raise extCode, , msg
'        End If
        
        Call DeleteDirectory(fileOrFoldrPath)
        
    ElseIf FileFound(fileOrFoldrPath) Then
        Call Kill(fileOrFoldrPath)
    End If
    
    ErrorIf F_.Found(fileOrFoldrPath) _
         , "Unable to delete the file/folder, but no error was raised."
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Delete", fileOrFoldrPath
End Sub

Public Sub DeleteContents(foldrPath As String)
    On Error GoTo Cleanup
    
    Call DeleteDirectory(foldrPath, deleteContntsOnly:=True)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "DeleteContents" ', "details of error"
End Sub


Public Function IsRemote(fileOrFoldrPath As String _
                       ) As Boolean
    On Error GoTo Cleanup
    
    If Left$(fileOrFoldrPath, 2) = "\\" _
        Then IsRemote = True
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "IsRemote" ', "details of error"
End Function


Public Property Get UserLocalTemp(Optional subFoldr As String _
                                ) As String
    Dim buffr$, retVal&
    On Error GoTo Cleanup
    
    buffr = String$(260, 0) ' MAX_PATH
    retVal = GetTempPath(Len(buffr), buffr)
    If retVal Then UserLocalTemp = Left$(buffr, retVal)
    
    If LenB(subFoldr) <> 0 Then _
        UserLocalTemp = UserLocalTemp & Folderize(subFoldr)
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "UserLocalTemp" ', "details of error"
End Property


Public Function MakeTempFile(Optional fileContnts As String _
                           , Optional prefx As String _
                           , Optional suffx As String _
                           , Optional extnsion As String = ".tmp" _
                           ) As String
    On Error GoTo Cleanup
    
    '  return path of temp file
    '
    MakeTempFile = TempFile(prefx, suffx, extnsion)
    
    Call CreateFile(MakeTempFile, fileContnts)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "MakeTempFile" ', "details of error"
End Function


Public Function ToString(pathOfFile As String _
                       ) As String
    Dim fn#
    On Error GoTo ErrH
    
    fn = FreeFile
    Open pathOfFile For Input As #fn
    ToString = StrConv(InputB(LOF(fn), fn), vbUnicode)
    Close #fn

ErrH: Blame "ToString"
End Function


Public Function TempFile(Optional prefx As String _
                       , Optional suffx As String _
                       , Optional extnsion As String = ".tmp" _
                       ) As String
    Static incrementr&
    On Error GoTo Cleanup
    
    TempFile = UserLocalTemp() & prefx & incrementr & suffx & extnsion
    
    If Found(TempFile) Then Delete TempFile
    
    incrementr = incrementr + 1
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "TempFile" ', "details of error"
End Function



Public Function RunCommand(lineOfCommand As String _
                         , Optional freezeCallingWindw As Boolean = False _
                         , Optional ByVal milliSecWaitTimeout As Long = -1 _
                         , Optional windwStyle As VbAppWinStyle = vbHide _
                         ) As Long
    On Error GoTo Cleanup
    
    Dim batchF$: batchF = F_.TempFile("RunCommand", , ".bat")
    
    Call F_.CreateFile(batchF, lineOfCommand)
    
    RunCommand = ShellWait(batchF, freezeCallingWindw _
                         , milliSecWaitTimeout, windwStyle)
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
            Call F_.Delete(batchF)
            
  LoadError "RunCommand" ', "details of error"
End Function



' *********************************************************************
'  Copyright ©1996-2005 Karl E. Peterson, All Rights Reserved
'  http://vb.mvps.org/samples
'  Shell32.Bas
' *********************************************************************
Public Function ShellWait(shellCommnd As String _
                        , Optional freezeCallingWindw As Boolean = False _
                        , Optional ByVal milliSecWaitTimeout As Long = -1 _
                        , Optional ByVal windwStyle As VbAppWinStyle = vbHide _
                        ) As Long
    Dim procssID&, hProcss&, fdwAccess&
    On Error GoTo Cleanup
    
    
    '  Validate parameters.
    '
    If (windwStyle < vbHide) Or (windwStyle > vbMinimizedNoFocus) _
        Then windwStyle = vbMinimizedNoFocus
    
    If milliSecWaitTimeout < -1 Then _
        milliSecWaitTimeout = -1 'INFINITE
    
    
    
    '  Launch process in requested mode.
    '
    procssID = Shell(shellCommnd, windwStyle)
    
    
    
    If freezeCallingWindw Then
        '
        '  Shells a new process and waits for it to complete.
        '   Calling application is totally non-responsive while new process executes.
        '
        fdwAccess = &H100000 'SYNCHRONIZE
        ' get handle to process, so we can wait for it to finish
        hProcss = OpenProcess(fdwAccess, False, procssID)
        If hProcss Then
            ShellWait = WaitForSingleObject(hProcss, milliSecWaitTimeout)
            Call CloseHandle(hProcss)
        End If
                                '  give feedback on what happened
        Select Case ShellWait      '
            Case &H102&:        Debug.Print "Timeout period elapsed"
            Case 0:             Debug.Print "Normal completion"
            Case &H80&:         Debug.Print "Wait Abandoned"
            Case -1&:           Debug.Print "Error on call: "; Err.LastDllError
        End Select
    
    
    
    Else ' Remain responsive
        '
        '  Shells a new process and waits for it to complete.
        '   Calling application is responsive while new process executes.
        '    It will react to new events,
        '     though execution of the current thread will not continue.
        '
        fdwAccess = &H400 'PROCESS_QUERY_INFORMATION
        ' get handle to process, so we can check for an exit code
        hProcss = OpenProcess(fdwAccess, False, procssID)
        If hProcss Then
            Do
                GetExitCodeProcess hProcss, ShellWait
                DoEvents
                Sleep 100
            Loop While ShellWait = &H103 'STILL_ACTIVE
            Call CloseHandle(hProcss)
        End If
    End If
    
        
Cleanup:        SaveError
                'Set someObj = Nothing
                'Erase someArray
      LoadError "ShellWait" ', "details of error"
End Function






Public Function List(ByVal foldrPath As String _
                   , Optional fileFiltr As String = "*.*" _
                   , Optional attrbute As VbFileAttribute = vbNormal + vbDirectory _
                   , Optional includPath As Boolean = False _
                   , Optional ByRef rsultCount As Long _
                   ) As String()
    Dim i&, dirResult$, t1D$()
    On Error GoTo Cleanup
    
    foldrPath = Folderize(foldrPath)
    rsultCount = 0
    
    dirResult = Dir$(foldrPath & fileFiltr, attrbute)
    
    Do While LenB(dirResult) <> 0
        
'        If attrbute = vbDirectory And _
'            Not IsFolder(foldrPath & dirResult) Then GoTo NextFile
        
        If dirResult = "." Or dirResult = ".." Then GoTo NextFile
            
        ReDim Preserve t1D(0 To rsultCount)
        
        For i = rsultCount To 1 Step -1
            ' look to see where in the list it belongs
            If UCase$(dirResult) >= UCase$(t1D(i - 1)) Then Exit For
            t1D(i) = t1D(i - 1)
        Next i
        
        t1D(i) = IIf(includPath, foldrPath, vbNullString) & dirResult
        
        
        '  if folder, append "\"
        '
        If IsFolder(foldrPath & dirResult) Then t1D(i) = t1D(i) & "\"
        
        rsultCount = rsultCount + 1
        
NextFile:   dirResult = Dir$() ' Get the next file.
    Loop
    
    List = t1D
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase t1D
  LoadError "List" ', "details of error"
End Function

Public Function Count(ByVal foldrPath As String _
                    , Optional fileFiltr As String = "*.*" _
                    , Optional attrbute As VbFileAttribute = vbNormal + vbDirectory _
                    ) As Long
    On Error GoTo Cleanup
    
    Call F_.List(foldrPath, fileFiltr _
               , attrbute, rsultCount:=Count)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Count" ', "details of error"
End Function

Public Function Folderize(foldrPath As String) As String
    Folderize = foldrPath _
              & IIf(Right$(foldrPath, 1) = "\", vbNullString, "\")
End Function






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub DeleteDirectory(ByVal foldrPath As String _
                          , Optional deleteContntsOnly As Boolean = False _
                          )
    Dim ff$(), i&, fCount&
    On Error GoTo Cleanup
    
    ' Get a list of files/folders it contains.
    ff = List(foldrPath, "*.*" _
            , vbReadOnly + vbHidden + vbSystem + vbDirectory _
            , True, fCount)
    
    
    '  delete each content
    '
    If fCount <> 0 Then
        For i = 0 To UBound(ff)
            
            '  if it's a directory,
            '   - call me again
            '
            If IsFolder(ff(i)) Then
                Call DeleteDirectory(ff(i))
            
            '  if normal file, do usual
            Else
                Call Delete(ff(i))
            End If
        Next i
    End If
    
    If Not deleteContntsOnly Then
        ' Remove the read-only flag if set.
        SetAttr foldrPath, vbNormal
        RmDir foldrPath
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ff
  LoadError "DeleteDirectory" ', "details of error"
End Sub

Private Function FileFound(pathOfFile As String) As Boolean
    If LenB(pathOfFile) <> 0 Then ' because Dir("") returns something
        On Error Resume Next
            FileFound = LenB(Dir(pathOfFile)) <> 0
        On Error GoTo 0
    End If
End Function

Public Function IsFolder(targetFoldrPath As String) As Boolean
    On Error Resume Next
        IsFolder = GetAttr(targetFoldrPath) And vbDirectory
    On Error GoTo 0
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As F_Errors = -1 _
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
                , Optional ByVal errorNumbr As F_Errors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As F_Errors = -1 _
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
