VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_MarkdownExe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITestCaseData

Private SUT As MarkdownExe
Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
Select Case Test.MethodName
    
    Case "ToHtml"
        Test.Use("aa").Expect "aa"
        Test.Use("*aa*").Expect "<em>aa</em>"
        Test.Use("**aa**").Expect "<strong>aa</strong>"
        'Test.Use("“á”").Expect "&ldquo;&aacute;&rdquo;"
End Select
End Sub


Public Function ToHtml(markDownTxt As String) As String
    ToHtml = SUT.ToHtml(markDownTxt)
    ToHtml = Mid$(ToHtml, 4, Len(ToHtml) - 7)
End Function




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Test Utility Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub DoSomething()
    
End Sub







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Fixture Framework Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub FixtureSetup()
    mMethodsRan = 0: mMethodCount = Sim.NewTestFixture(Me).TestCount
    
End Sub


Public Sub Setup()
    Set SUT = New MarkdownExe
    Call SUT.Init("C:\Toolbox\Markdown\multimarkdown.exe", parsr_MultiMarkdown_1)
End Sub


Public Sub Teardown()
    Set SUT = Nothing
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    'TestBed.QuitExcel
End Sub
