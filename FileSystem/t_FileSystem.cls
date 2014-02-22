VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_FileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseData

Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
Dim r$, f$
Select Case Test.MethodName
    
    Case "CreateFolder_ParentExists"
        r = App.Path: f = "newFolder": Test.Use(r, f).TestName "App.Path \ " & f
        r = App.Path: f = "newFolder\": Test.Use(r, f).TestName "App.Path \ " & f
        r = "C:": f = "newFolder": Test.Use(r, f).TestName r & " \ " & f
        r = "C:": f = "newFolder\": Test.Use(r, f).TestName r & " \ " & f
    
    Case "CreateFolder_AlreadyCreated"
        r = App.Path: f = "newFolder": Test.Use(r, f).TestName "App.Path \ " & f
        r = App.Path: f = "newFolder\": Test.Use(r, f).TestName "App.Path \ " & f
        r = "C:": f = "newFolder": Test.Use(r, f).TestName r & " \ " & f
        r = "C:": f = "newFolder\": Test.Use(r, f).TestName r & " \ " & f
    
    Case "CreateFolder_Nested"
        r = App.Path: Test.Use(r, "parent", "target").TestName "App.Path - 1 level - no tail"
        r = App.Path: Test.Use(r, "parent", "target\").TestName "App.Path - 1 level - with tail"
        r = App.Path: Test.Use(r, "p1", "p2", "target").TestName "App.Path - 2 levels - no tail"
        r = App.Path: Test.Use(r, "p1", "p2", "target\").TestName "App.Path - 2 levels - with tail"
        r = App.Path: Test.Use(r, "p1", "p2", "p3", "p4", "target").TestName "App.Path - 4 levels"
        
        r = "C:": Test.Use(r, "parent", "target").TestName r & " - 1 level - no tail"
        r = "C:": Test.Use(r, "parent", "target\").TestName r & " - 1 level - with tail"
        r = "C:": Test.Use(r, "p1", "p2", "target").TestName r & " - 2 levels - no tail"
        r = "C:": Test.Use(r, "p1", "p2", "target\").TestName r & " - 2 levels - with tail"
        r = "C:": Test.Use(r, "p1", "p2", "p3", "p4", "target").TestName r & " - 4 levels"
        
        r = "\\Vboxsvr\Shared": Test.Use(r, "parent", "target").TestName r & " - 1 level - no tail"
        r = "\\Vboxsvr\Shared": Test.Use(r, "parent", "target\").TestName r & " - 1 level - with tail"
        r = "\\Vboxsvr\Shared": Test.Use(r, "p1", "p2", "target").TestName r & " - 2 levels - no tail"
        r = "\\Vboxsvr\Shared": Test.Use(r, "p1", "p2", "target\").TestName r & " - 2 levels - with tail"
        r = "\\Vboxsvr\Shared": Test.Use(r, "p1", "p2", "p3", "p4", "target").TestName r & " - 4 levels"
        
        r = "F:": Test.Use(r, "parent", "target").TestName r & " - 1 level - no tail"
        r = "F:": Test.Use(r, "parent", "target\").TestName r & " - 1 level - with tail"
        r = "F:": Test.Use(r, "p1", "p2", "target").TestName r & " - 2 levels - no tail"
        r = "F:": Test.Use(r, "p1", "p2", "target\").TestName r & " - 2 levels - with tail"
        r = "F:": Test.Use(r, "p1", "p2", "p3", "p4", "target").TestName r & " - 4 levels"
        
    
    Case "Delete_ErasesFolder"
        Test.Use("").TestName "empty foldr"
        Test.Use("subfolder\").TestName "1 subfolder"
        Test.Use("sample.file").TestName "1 file"
        Test.Use("folder1\", "folder2\", "file.txt").TestName "2 subfolders, 1 file"
        
        
    Case "Parent"
        r = "C:": Test.Use(r).Expect("").TestName r
        r = "C:\": Test.Use(r).Expect("").TestName r
        r = "C:\a": Test.Use(r).Expect("C:\").TestName r
        r = "C:\a\": Test.Use(r).Expect("C:\").TestName r
        r = "C:\a\b": Test.Use(r).Expect("C:\a\").TestName r
        r = "C:\a\b\": Test.Use(r).Expect("C:\a\").TestName r
        r = "C:\a\b\c": Test.Use(r).Expect("C:\a\b\").TestName r
        r = "C:\a\b\c\": Test.Use(r).Expect("C:\a\b\").TestName r
        
        r = "C:\f.txt": Test.Use(r).Expect("C:\").TestName r
        r = "C:\a\f.txt": Test.Use(r).Expect("C:\a\").TestName r
        r = "C:\a\b\f.txt": Test.Use(r).Expect("C:\a\b\").TestName r
        r = "C:\a\b\c\f.txt": Test.Use(r).Expect("C:\a\b\c\").TestName r
        
        
    
    Case "FileName"
        r = "file.txt": Test.Use(r, True).Expect("file.txt").TestName r & " - with ext"
        r = "C:\file.txt": Test.Use(r, True).Expect("file.txt").TestName r & " - with ext"
        r = "C:\a\file.txt": Test.Use(r, True).Expect("file.txt").TestName r & " - with ext"
        r = "C:\a\b\file.txt": Test.Use(r, True).Expect("file.txt").TestName r & " - with ext"
        r = "C:\a\b\c\file.txt": Test.Use(r, True).Expect("file.txt").TestName r & " - with ext"
        r = "\\a\b\c\file.txt": Test.Use(r, True).Expect("file.txt").TestName r & " - with ext"
        r = "C:\a\b\c\folder\": Test.Use(r, True).Expect("").TestName r & " - with ext"
        
        r = "file.txt": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\file.txt": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\a\file.txt": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\a\b\file.txt": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\a\b\c\file.txt": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "\\a\b\c\file.txt": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\a\b\c\folder\": Test.Use(r, False).Expect("").TestName r & " - no ext"
        
        r = "file": Test.Use(r, True).Expect("file").TestName r & " - with ext"
        r = "file": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\file": Test.Use(r, True).Expect("file").TestName r & " - with ext"
        r = "C:\file": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        r = "C:\a\b\file": Test.Use(r, True).Expect("file").TestName r & " - with ext"
        r = "C:\a\b\file": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        
        r = "C:\a\file.txt.zip": Test.Use(r, True).Expect("file.txt.zip").TestName r & " - with ext"
        r = "C:\a\file.txt.zip": Test.Use(r, False).Expect("file").TestName r & " - no ext"
        
    
    Case "CopyFileToFolder"
        Test.Use("target", True).TestName "no slash, existing"
        Test.Use("target\", True).TestName "with slash, existing"
        Test.Use("target\", False).TestName "with slash, non-existent"
    
    
    Case "CopyFileToFile"
        r = "newFile.txt": Test.Use(r, False).TestName r & " : non-existent"
        r = "folder\newFile.txt": Test.Use(r, False).TestName r & " : non-existent"
        
        r = "newFile.txt": Test.Use(r, True).TestName r & " : existing"
        r = "newFile": Test.Use(r, True).TestName r & " : existing"
    
    
    Case "List_All"
        Test.Use("file.txt").TestName "1 file"
        Test.Use("folder\").TestName "1 folder"
        Test.Use("f1.txt", "f2.txt").TestName "2 files"
        Test.Use("f1\", "f2\").TestName " 2 folders"
        Test.Use("f1.txt", "f2.txt", "f3.txt", "f4.txt").TestName "4 files"
        Test.Use("f1\", "f2\", "f3\", "f4\").TestName "4 folders"
        Test.Use("file.txt", "folder\").TestName "1 file, 1 folder"
        Test.Use("f1\", "f2\", "file.txt").TestName "1 file, 2 folders"
        Test.Use("f1.txt", "f2.txt", "folder\").TestName "2 files, 1 folder"
    
    
    Case "List_Filtered"
        Test.Use("*.opf", vbNormal _
               , "a.txt" _
               , "bdir\" _
               , "c.jpg" _
               , "d.opf" _
               , "e.gif" _
               ).Expect ss("d.opf")
    
        Test.Use("*.opf", vbNormal _
               , "a.opf" _
               , "bdir\" _
               , "c.jpg" _
               , "d.opf" _
               , "e.gif" _
               ).Expect ss("a.opf", "d.opf")
    
'        Test.Use("*.*", vbDirectory _
'               , "a.txt" _
'               , "bdir\" _
'               , "c.jpg" _
'               , "d.opf" _
'               , "e.gif" _
'               ).Expect ss("bdir")
    
End Select
End Sub


Public Function List_Filtered(filtr As String _
                            , attrbute As VbFileAttribute _
                            , ParamArray contnts() As Variant _
                            ) As String()
    Dim tmp$, i&, fCount&
    
    ' create test files here
    tmp = F_.UserLocalTemp("TestCopy")
    Call CreateInFolder(tmp, contnts)
    
    
    ' return filtered list
    List_Filtered = F_.List(tmp, filtr, attrbute)
    
    
    ' Cleanup test context.
     F_.Delete tmp
End Function


Public Sub List_All(ParamArray contnts() As Variant)
    Dim tmp$, i&, fCount&
    
    ' create test files here
    tmp = F_.UserLocalTemp("TestCopy")
    Call CreateInFolder(tmp, contnts)
    
    
    ' Execute then verify
    Assert.That F_.List(tmp, rsultCount:=fCount) _
              , Iz.EqualTo(contnts).AsCollection
    
    Assert.That fCount, Iz.EqualTo(UBound(contnts) + 1)
    
    ' Cleanup test context.
     F_.Delete tmp
End Sub



'    Case "CopyFileToExistingFile"
'    Case "CopyFileToNonExistentFile"
'          CopyFileToSameFolder

Public Sub CopyFileToFolder(ByVal targetFoldr As String _
                          , createTargetFoldr As Boolean)
    Dim tmp$, src$, expctd$
    
    ' manipulate test files here
    tmp = F_.UserLocalTemp("TestCopy")
    
    
    ' Create temp source file
    src = F_.CreateFile(F_.TempFile())
    Assert.That src, Path.Exists
    
    
    targetFoldr = tmp & targetFoldr
    
    ' Create target folder as needed
    If createTargetFoldr Then
        F_.CreateFolder targetFoldr
        Assert.That targetFoldr, Path.Exists
    End If
    
    
    ' Execute method under test.
    Call F_.Copy(src, targetFoldr)
    
    
    ' Verify result.
    expctd = Folderize(targetFoldr) & FileName(src)
    Assert.That expctd, Path.Exists
    
    
    ' Cleanup test context.
    F_.Delete tmp
End Sub


Public Sub CopyFileToFile(targtFilename As String _
                        , createTargtFile As String)
    Dim tmp$, src$, targtPath$
    
    ' manipulate test files here
    tmp = F_.UserLocalTemp("TestCopy")
    
    
    ' Create temp source file
    src = F_.CreateFile(F_.TempFile())
    Assert.That src, Path.Exists
    
    
    ' expect target to be in temp folder
    targtPath = tmp & targtFilename
    
    
    ' create target file as needed
    If createTargtFile Then
        F_.CreateFile targtPath
        Assert.That targtPath, Path.Exists
    End If
    
    
    ' Execute method under test.
    Call F_.Copy(src, targtPath)
    
    
    ' Verify result.
    Assert.That targtPath, Path.Exists
    
    
    ' Cleanup test context.
    F_.Delete tmp
End Sub



Public Function FileName(pathOfFile As String _
                       , Optional includeExtnsion As Boolean = True _
                       ) As String
    FileName = F_.FileName(pathOfFile, includeExtnsion)
End Function


Public Function Parent(fileOrFoldrPath As String _
                     ) As String
    Parent = F_.Parent(fileOrFoldrPath)
End Function


Public Sub UserLocalTemp_SameAsFSO()
    Dim actualVal$, expctd$
    
    expctd = New_RC4.FSO.GetTmpPath
    
    actualVal = F_.UserLocalTemp
    
    Assert.That actualVal, Iz.EqualTo(expctd)
    
End Sub



Public Sub Delete_ErasesFolder(ParamArray foldrContnts() As Variant)
    
    '  context req:
    '   - folder to be deleted must not exist
    '
    Dim foldr$: foldr = F_.UserLocalTemp("delete_test")
    Call F_.Delete(foldr)
    Assert.That foldr, Path.DoesNotExist
    
    
    '  create the folder to be deleted
    '
    Call F_.CreateFolder(foldr)
    
        
    '  context req:
    '   - folder should have contents
    '
    Dim i&, contnt$
    For i = 0 To UBound(foldrContnts)
        
        contnt = foldr & foldrContnts(i)
        
        ' if has "\", create subfolder
        If S_(CStr(foldrContnts(i))).Has("\") Then
            Call F_.CreateFolder(contnt)
        
        ' if has ".", create blank file
        ElseIf S_(CStr(foldrContnts(i))).Has(".") Then
            Call F_.CreateFile(contnt)
        End If
        
        ' verify creation
        Assert.That contnt, Path.Exists
    Next i
    
    
    ' Execute method under test.
    Call F_.Delete(foldr)
    
    ' Verify result.
    Assert.That foldr, Path.DoesNotExist
    
End Sub



Public Sub CreateFolder_ParentExists(rootFoldr$, foldrName$)
    
    '  context req:
    '   - parent exists
    '
    Dim f$: f = rootFoldr & "\" & foldrName
    
    
    '  context req:
    '   - folder not yet created
    '
    Call F_.Delete(f)
    Assert.That f, Path.DoesNotExist
    
    
    '  execute, verify, then cleanup
    '
    Call F_.CreateFolder(f)
    Assert.That f, Path.Exists
    Call F_.Delete(f)
End Sub


Public Sub CreateFolder_AlreadyCreated(rootFoldr$, foldrName$)
    
    '  context req:
    '   - parent exists
    '
    Dim f$: f = rootFoldr & "\" & foldrName
    
    
    '  context req:
    '   - folder already created
    '
    Call F_.CreateFolder(f)
    Assert.That f, Path.Exists
    
    
    '  execute, verify, then cleanup
    '
    Call F_.CreateFolder(f)
    Assert.That f, Path.Exists
    Call F_.Delete(f)
End Sub


Public Sub CreateFolder_Nested(ParamArray foldrs() As Variant)
    Dim targt$, parnt$, i&, ss$()
    
    '  assemble parent
    '   - Join() all folders except last
    '
    ReDim ss(UBound(foldrs) - 1)
    For i = 0 To UBound(ss)
        ss(i) = foldrs(i)
    Next i
    parnt = Join(ss, "\")
    Assert.That parnt, Path.DoesNotExist
    
    
    '  create inside non-existent parent
    '
    targt = parnt & "\" & foldrs(UBound(foldrs))
    
    
    '  execute then verify
    '
    Call F_.CreateFolder(targt)
    Assert.That targt, Path.Exists
    
    
    '  cleanup
    '
    For i = UBound(foldrs) To 1 Step -1
        Call F_.Delete(targt)
        targt = Left$(targt, Len(targt) - Len(foldrs(i)) - 1)
    Next i
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Test Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub CreateInFolder(ByVal targetFoldr As String _
                         , ParamArray contnts() As Variant)
    
    Dim ff: ff = contnts(0)
    targetFoldr = Folderize(targetFoldr)
    
    
    '  delete contents if any
    '
    F_.DeleteContents targetFoldr
    Assert.That F_.Count(targetFoldr), Iz.EqualTo(0) _
             , "Folder should be emptied."
    
    
    Dim i&: For i = 0 To UBound(ff)
        If Right(ff(i), 1) = "\" Then
            F_.CreateFolder targetFoldr & ff(i)
        Else
            F_.CreateFile targetFoldr & ff(i)
        End If
    Next i
    
    
    '  verify created files/folders
    '
    Assert.That F_.List(targetFoldr), Iz.EqualTo(ff).AsCollection _
             , "All files/folders in param should be created."
End Sub

Private Property Get S_(strText As String) As StringWrapper
    Static obj As StringWrapper
    
    If obj Is Nothing Then _
        Set obj = New StringWrapper
    
    Call obj.Init(strText)
    Set S_ = obj
    
End Property

'  returns String() array containing params
'
Private Function ss(ParamArray arrayItms() As Variant _
                  ) As String()
    Dim s$(): ReDim s(UBound(arrayItms))
    Dim i&: For i = 0 To UBound(s)
        s(i) = CStr(arrayItms(i))
    Next i
    ss = s
    Erase s
End Function




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Fixture Framework Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub FixtureSetup()
    mMethodsRan = 0: mMethodCount = Sim.NewTestFixture(Me).TestCount
    
End Sub


Public Sub Setup()
    'Set SUT = New
    'Call SUT.Init(
End Sub


Public Sub Teardown()
    'Set SUT = Nothing
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    'TestBed.QuitExcel
End Sub
