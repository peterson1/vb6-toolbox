Attribute VB_Name = "RC4Factory"
Option Explicit
Const MODULE_NAME$ = "RC4Factory"

Public Const vbRichClient4_dll$ = "vbRichClient4.dll"
Public Const vb_cairo_sqlite_dll$ = "vb_cairo_sqlite.dll"
Public Const DirectCOM_dll$ = "DirectCOM.dll"

Private Enum RC4FactoryErrors       ' you may make this Public for tests
    ErrorBase = 3000                ' you may adjust this minimum
    PathNotFound
    NotInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As RC4FactoryErrors      '
    Description As String
End Type
Private mError As ErrorHolder


Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" _
                  (ByVal LibFilePath As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" _
                  (ByVal hLibModule As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" _
                  (ByVal DllName As Long) As Long
                  
Private Declare Function GetInstanceEx Lib "DirectCOM.dll" _
                  (ByRef StrPtr_FilePath As Long, _
                   ByRef StrPtr_ClassName As Long, _
          Optional ByVal UseAlteredSearchPath As Boolean = True) As Object


Private mInitted As Boolean _
      , mUseRegFree As Boolean

Private mFactory As vbRichClient4.cFactory










' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Factories
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get New_RC4() As vbRichClient4.cConstructor
    On Error GoTo Cleanup
    
    If InIDE Then
        Call RC4Factory.Init(App.Path)
    
    Else
        ErrorIf Not mInitted, "", NotInitted
    End If
    
    Set New_RC4 = mFactory.c
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Get]New_RC4" ', "Line: " & Erl
End Property

Public Property Get New_RC4RegFree() As vbRichClient4.cRegFree
    On Error GoTo Cleanup
    
    ErrorIf Not mInitted, "", NotInitted
    
    Set New_RC4RegFree = mFactory.regfree
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "[Get]New_RC4RegFree" ', "Line: " & Erl
End Property

'Public Property Get TestRC4() As vbRichClient4.cConstructor
'    Static obj As vbRichClient4.cFactory
'
'    ErrorIf Not InIDE _
'         , "You should only use this method within the IDE."
'
'    If obj Is Nothing Then _
'        Set obj = New vbRichClient4.cFactory
'
'    Set TestRC4 = obj.c
'
'End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(vbRichClient4DllsFoldrPath As String _
              , Optional forceRegFree As Boolean = False)
    Dim foldr$
    On Error GoTo Cleanup
    
    
    '  do not proceed if already Iniited
    '
    If mInitted Then Exit Sub
    
    
    
    '  use Reg-Free if NOT running in IDE (compiled/deployed)
    '   - or if forced to do so
    '
    mUseRegFree = forceRegFree Or Not InIDE
    If mUseRegFree Then
        
        
        '  look for required files/folder
        '
        foldr = Folderize(vbRichClient4DllsFoldrPath)
        ErrorIfMissing foldr, "vbRichClient DLLs folder"
        ErrorIfMissing foldr & vbRichClient4_dll, "vbRichClient4 Dll"
        ErrorIfMissing foldr & vb_cairo_sqlite_dll, "Cairo-SQLite Dll"
        ErrorIfMissing foldr & DirectCOM_dll, "Direct COM Dll"
        
        
        '  load library for DirectCOM.dll
        '
        Call LoadLibrary(StrPtr(foldr & DirectCOM_dll))
        
        
        '  using reg-free DirectCOM
        '   - get Factory instance from vbRichClient.dll
        '
        Set mFactory = GetInstanceEx( _
                            StrPtr(foldr & vbRichClient4_dll) _
                          , StrPtr("cFactory"))
    
    
    '  otherwise,
    '   - use normal instantiation - from Registry
    '
    Else
        Set mFactory = New vbRichClient4.cFactory
    End If
    
    
    '  remember Initted state so we don't repeat
    '
    mInitted = True
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Init" ', "details of error"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Deconstructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Cleanup()
    On Error GoTo Cleanup
    
    mInitted = False
    
    If Not mFactory Is Nothing Then _
        mFactory.c.CleanupRichClientDll
    
    Set mFactory = Nothing
    
    If mUseRegFree Then _
        Call UnloadDlls(vbRichClient4_dll _
                      , vb_cairo_sqlite_dll _
                      , DirectCOM_dll)
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Cleanup" ', "Line: " & Erl
End Sub








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Public Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Function WellFormedXML(xmlStr As String) As Boolean
    WellFormedXML = New_RC4.SimpleDOM(xmlStr).WellFormed
End Function

Public Function SHA1(strText As String) As String
    SHA1 = New_RC4.Crypt.SHA1(strText, True)
End Function

Public Function CairoQuote(quotText As String _
                         , quotCaption As String _
                         , maxmumWidth As Double _
                         , maxmumHeight As Double _
                         ) As cCairoSurface
    Dim nextY#, c As cCairoContext
    On Error GoTo Cleanup
    DoEvents
    
    Const bgColr& = -2147483633
    Const fntName$ = "Times New Roman"  ' "MS Serif" '
    Const fntSize& = 10
    Dim fntColr&: fntColr = RGB(140, 140, 140)
    
    Set CairoQuote = New_RC4.Cairo.CreateSurface(maxmumWidth, maxmumHeight)
    Set c = CairoQuote.CreateContext
    
    
    '  draw background
    '
    c.SetSourceColor bgColr: c.Paint
    
    
    '  draw quote text
    '
    c.SelectFont fntName, fntSize, fntColr
    nextY = DrawText(quotText, c, 0, 0, maxmumWidth, vbLeftJustify)
    
    
    '  draw quote caption
    '
    c.SelectFont fntName, fntSize, fntColr
    nextY = DrawText(quotCaption, c, 0, nextY + 7 _
                   , maxmumWidth, vbRightJustify)
    
    
    
Cleanup:    SaveError
            Set c = Nothing
            'Erase someArray
  LoadError "CairoQuote" ', "details of error"
End Function

Private Function DrawText(strText As String _
                        , cairoContxt As cCairoContext _
                        , xPosition As Double _
                        , yPosition As Double _
                        , maxmumWidth As Double _
                        , textAlignmnt As AlignmentConstants _
                        ) As Double
    Dim numRows&, extntH#, redrawAttmpts&
    On Error GoTo Cleanup
    With cairoContxt
        
        
        '  calculate height of text-extent
        '
        numRows = .DrawText(0, 0, maxmumWidth, 0, strText, , textAlignmnt)
        extntH = numRows * .GetFontHeight()
        
        
        '  render actual text
        '
RedrawText:
        .DrawText xPosition, yPosition, maxmumWidth, extntH, strText, , textAlignmnt
        
        
        '  return next clear yPosition
        '
        DrawText = yPosition + extntH
        
    End With
Cleanup:    If Err = 16 And _
               redrawAttmpts < 1 Then
                redrawAttmpts = redrawAttmpts + 1
                Err.Clear
                GoTo RedrawText
            End If
            SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "DrawText" ', "details of error"
End Function








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function Folderize(foldrPath As String) As String
    Folderize = foldrPath _
              & IIf(Right$(foldrPath, 1) = "\", vbNullString, "\")
End Function

Private Function FileFound(pathOfFile As String) As Boolean
    If LenB(pathOfFile) <> 0 Then ' because Dir("") returns something
        On Error Resume Next
            FileFound = LenB(Dir(pathOfFile)) <> 0
        On Error GoTo 0
    End If
End Function

Private Function FolderFound(targetFoldrPath As String) As Boolean
    On Error Resume Next
        FolderFound = GetAttr(targetFoldrPath) And vbDirectory
    On Error GoTo 0
End Function

Private Sub UnloadDlls(ParamArray namesOfDlls() As Variant)
    Dim i&, fName$, moduleHandl&
    On Error GoTo Cleanup
    
    For i = 0 To UBound(namesOfDlls)
        fName = namesOfDlls(i)
        moduleHandl = GetModuleHandle(StrPtr(fName))
        
        While moduleHandl <> 0
            FreeLibrary moduleHandl
            moduleHandl = GetModuleHandle(StrPtr(fName))
        Wend
        
    Next i
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "UnloadDlls" ', "Line: " & Erl
End Sub

Public Function InIDE() As Boolean
    On Error Resume Next
        Debug.Print 1 / 0
        InIDE = Err.Number
    Err.Clear
End Function

Private Sub ErrorIfMissing(fileOrFoldrPath As String _
                         , Optional fileOrFoldrDescrption As String = "File" _
                         , Optional errorNumbr As RC4FactoryErrors = PathNotFound _
                         )
    ErrorIf LenB(fileOrFoldrPath) = 0 _
          , fileOrFoldrDescrption & " not found. - path is [blank]", errorNumbr
    
    If FileFound(fileOrFoldrPath) Then Exit Sub
    If FolderFound(fileOrFoldrPath) Then Exit Sub
    
    Err.Raise errorNumbr, MODULE_NAME _
            , fileOrFoldrDescrption & " not found:" & vbCrLf & fileOrFoldrPath
End Sub







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As RC4FactoryErrors = -1 _
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
                    , Optional ByVal errorNumbr As RC4FactoryErrors = -1 _
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

