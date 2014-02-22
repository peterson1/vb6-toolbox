VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_StringWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private SUT As StringWrapper
Private mMethodsRan&, mMethodCount&





Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
Select Case Test.MethodName
    
    Case "Has_GetBoolean"
          Test.Use("abcd", "a", True).Expect (True)
          Test.Use("abcd", "b", True).Expect (True)
          Test.Use("abcd", "c", True).Expect (True)
          Test.Use("abcd", "d", True).Expect (True)
         Test.Use("abcd", "e", True).Expect (False)
         Test.Use("abcd", "a", False).Expect (True)
         Test.Use("abcd", "b", False).Expect (True)
         Test.Use("abcd", "c", False).Expect (True)
         Test.Use("abcd", "d", False).Expect (True)
        Test.Use("abcd", "e", False).Expect (False)

         Test.Use("ABCD", "a", True).Expect (False)
         Test.Use("ABCD", "b", True).Expect (False)
         Test.Use("ABCD", "c", True).Expect (False)
         Test.Use("ABCD", "d", True).Expect (False)
         Test.Use("ABCD", "e", True).Expect (False)
         Test.Use("ABCD", "a", False).Expect (True)
         Test.Use("ABCD", "b", False).Expect (True)
         Test.Use("ABCD", "c", False).Expect (True)
         Test.Use("ABCD", "d", False).Expect (True)
        Test.Use("ABCD", "e", False).Expect (False)

         Test.Use("abcd", "A", True).Expect (False)
         Test.Use("abcd", "B", True).Expect (False)
         Test.Use("abcd", "C", True).Expect (False)
         Test.Use("abcd", "D", True).Expect (False)
         Test.Use("abcd", "E", True).Expect (False)
         Test.Use("abcd", "A", False).Expect (True)
         Test.Use("abcd", "B", False).Expect (True)
         Test.Use("abcd", "C", False).Expect (True)
         Test.Use("abcd", "D", False).Expect (True)
        Test.Use("abcd", "E", False).Expect (False)

          Test.Use("ABCD", "A", True).Expect (True)
          Test.Use("ABCD", "B", True).Expect (True)
          Test.Use("ABCD", "C", True).Expect (True)
          Test.Use("ABCD", "D", True).Expect (True)
         Test.Use("ABCD", "E", True).Expect (False)
         Test.Use("ABCD", "A", False).Expect (True)
         Test.Use("ABCD", "B", False).Expect (True)
         Test.Use("ABCD", "C", False).Expect (True)
         Test.Use("ABCD", "D", False).Expect (True)
        Test.Use("ABCD", "E", False).Expect (False)
        
    
    Case "Has_GetPosition"
         Test.Use("abcd", "a", True).Expect (1)
         Test.Use("abcd", "b", True).Expect (2)
         Test.Use("abcd", "c", True).Expect (3)
         Test.Use("abcd", "d", True).Expect (4)
         Test.Use("abcd", "e", True).Expect (0)
        Test.Use("abcd", "a", False).Expect (1)
        Test.Use("abcd", "b", False).Expect (2)
        Test.Use("abcd", "c", False).Expect (3)
        Test.Use("abcd", "d", False).Expect (4)
        Test.Use("abcd", "e", False).Expect (0)

         Test.Use("ABCD", "a", True).Expect (0)
         Test.Use("ABCD", "b", True).Expect (0)
         Test.Use("ABCD", "c", True).Expect (0)
         Test.Use("ABCD", "d", True).Expect (0)
         Test.Use("ABCD", "e", True).Expect (0)
        Test.Use("ABCD", "a", False).Expect (1)
        Test.Use("ABCD", "b", False).Expect (2)
        Test.Use("ABCD", "c", False).Expect (3)
        Test.Use("ABCD", "d", False).Expect (4)
        Test.Use("ABCD", "e", False).Expect (0)

         Test.Use("abcd", "A", True).Expect (0)
         Test.Use("abcd", "B", True).Expect (0)
         Test.Use("abcd", "C", True).Expect (0)
         Test.Use("abcd", "D", True).Expect (0)
         Test.Use("abcd", "E", True).Expect (0)
        Test.Use("abcd", "A", False).Expect (1)
        Test.Use("abcd", "B", False).Expect (2)
        Test.Use("abcd", "C", False).Expect (3)
        Test.Use("abcd", "D", False).Expect (4)
        Test.Use("abcd", "E", False).Expect (0)

         Test.Use("ABCD", "A", True).Expect (1)
         Test.Use("ABCD", "B", True).Expect (2)
         Test.Use("ABCD", "C", True).Expect (3)
         Test.Use("ABCD", "D", True).Expect (4)
         Test.Use("ABCD", "E", True).Expect (0)
        Test.Use("ABCD", "A", False).Expect (1)
        Test.Use("ABCD", "B", False).Expect (2)
        Test.Use("ABCD", "C", False).Expect (3)
        Test.Use("ABCD", "D", False).Expect (4)
        Test.Use("ABCD", "E", False).Expect (0)
    
End Select
End Sub


Public Function Has_GetBoolean(fullStrng As String _
                             , subStrng As String _
                             , caseSensitiv As Boolean _
                             ) As Boolean
    
    '  Execute method under test.
    '   Then return result.
    '
    Has_GetBoolean = S_(fullStrng).Has(subStrng, caseSensitiv)
    
End Function


Public Function Has_GetPosition(fullStrng As String _
                              , subStrng As String _
                              , caseSensitiv As Boolean _
                              ) As Long
    
    '  Execute method under test.
    '   Then return result.
    '
    Has_GetPosition = S_(fullStrng).Has(subStrng, caseSensitiv)
    
End Function






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Test Utility Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Property Get S_(strText As String) As StringWrapper
    Call SUT.Init(strText)
    Set S_ = SUT
End Property







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Fixture Framework Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub FixtureSetup()
    mMethodsRan = 0: mMethodCount = Sim.NewTestFixture(Me).TestCount
    
    Set SUT = New StringWrapper
    
End Sub


Public Sub Setup()
    '
End Sub


Public Sub Teardown()
    '
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    Set SUT = Nothing
    'TestBed.QuitExcel
End Sub
