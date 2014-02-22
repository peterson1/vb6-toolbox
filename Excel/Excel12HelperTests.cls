VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Excel12HelperTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private SUT As Excel12Helper
Private mSampleTree As TreeSampler

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        
        Case "ColumnNumber"
            Test.Use("a").Expect 1
            Test.Use("A").Expect 1
            Test.Use("b").Expect 2
            Test.Use("B").Expect 2
            Test.Use("z").Expect 26
            Test.Use("Z").Expect 26
            'Test.Use("AA").Expect 27
        
        Case "ColumnLetter"
            Test.Use(1).Expect "A"
            Test.Use(2).Expect "B"
            Test.Use(25).Expect "Y"
            Test.Use(26).Expect "Z"
            Test.Use(27).Expect "AA"
            Test.Use(28).Expect "AB"
            Test.Use(128).Expect "DX"
        
        Case "CompliantRangeName"
            Test.Use("no space allowed").Expect "no_space_allowed"
            Test.Use("UPPER or lower case ReTaInEd").Expect "UPPER_or_lower_case_ReTaInEd"
            Test.Use("_underscores_ are allowed").Expect "_underscores_are_allowed"
            Test.Use(".dots. are allowed.").Expect ".dots._are_allowed."
            Test.Use("BS Elim Inv_Eqty").Expect "BS_Elim_Inv_Eqty"
            Test.Use("with [square] brackets").Expect "with_square_brackets"
            Test.Use("with (parenthesis) and: colon").Expect "with_parenthesis_and_colon"
        
    End Select
End Sub


Public Sub DrawTreeCollection()
    Dim treeModl As ITreeNode, t2D(), i&, wrkbook As Workbook
    Set treeModl = mSampleTree.AsTreeNodes(tr_CreatureNames)
    
    With New Excel.Application
'        With .Workbooks.Open("C:\Dropbox\Aldrus\Lorrie\etc\Schema-AINA Conso.xlsx", , True)
'            With .Worksheets("SCOA")
'                t2D = .Range(.Range("A3").End(xlDown) _
'                           , .Range("A3").End(xlToRight)).Value
'            End With
'            .Close False
'        End With
'        Set treeModl = New TreeBranch
'        Call treeModl.FromArray2D(t2D)
        
        Set wrkbook = .Workbooks.Add
        
        With treeModl
            For i = 0 To .Count - 1
                SUT.DrawTreeNodes .Node(i) _
                                 , wrkbook.Worksheets.Add, "A", 1
            Next i
        End With
        
        wrkbook.Application.Visible = True
        MsgBox Rand.mPearl
        wrkbook.Close False
    End With
    
    
End Sub




Public Function CompliantRangeName(strText$) As String
    CompliantRangeName = SUT.CompliantRangeName(strText)
End Function


Public Function ColumnLetter(columnIndx&) As String
    ColumnLetter = SUT.ColumnLetter(columnIndx)
End Function


Public Function ColumnNumber(columnLettr$) As Long
    ColumnNumber = SUT.ColumnNumber(columnLettr)
End Function


Public Sub Setup()
    Set mSampleTree = New TreeSampler
    Set SUT = New Excel12Helper
End Sub


Public Sub Teardown()
    Set SUT = Nothing
    Set mSampleTree = Nothing
End Sub
