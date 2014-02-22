VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseData

Private SUT1 As wgtTextField
      
Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
    Select Case Test.MethodName
        
        Case "TrimRight"
            Test.Use("tanggal isa", 1).Expect("tanggal is").TestName "Trim 1 char"
            Test.Use("tanggal adwa", 2).Expect("tanggal ad").TestName "Trim 2 chars"
            Test.Use("tanggal apat", 4).Expect("tanggal ").TestName "Trim 4 chars"
            
            Test.Use("tanggal wala", 0).Expect("tanggal wala").TestName "Trim 0 chars"
            Test.Use("", 1).Expect("").TestName "Trim blank string"
            
            Test.Use("sobra", 10).Expect("sobra").TestName "Trim too much"
            Test.Use("lahat", 5).Expect("").TestName "Trim all"
        
    End Select
End Sub



Public Function TrimRight(origText As String _
                       , Optional charsToTrim As Long = 1 _
                       ) As String
    
    TrimRight = SUT1.TrimRight(origText, charsToTrim)
    
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
    Set SUT1 = New wgtTextField
End Sub


Public Sub Teardown()
    Set SUT1 = Nothing
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    'TestBed.QuitExcel
End Sub
