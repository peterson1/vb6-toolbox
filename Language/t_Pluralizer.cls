VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_Pluralizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
Select Case Test.MethodName
    
    Case "Plu_ral"
        Test.Use("apple").Expect "apples"
        Test.Use("apple", -1).Expect "apples"
        Test.Use("apple", 0).Expect "No apples"
        Test.Use("apple", 1).Expect "1 apple"
        Test.Use("apple", 2).Expect "2 apples"
        Test.Use("apple", 1234).Expect "1,234 apples"
        Test.Use("apple", 1234567).Expect "1,234,567 apples"
        Test.Use("Apple", 9).Expect "9 Apples"
        Test.Use("APPLE", 4567).Expect "4,567 APPLES"
        
        
        '  test ~y
        '
        Test.Use("Category").Expect "Categories"
        Test.Use("category").Expect "categories"
        Test.Use("CATEGORY").Expect "CATEGORIES"
        Test.Use("agency").Expect "agencies"
        Test.Use("Agency").Expect "Agencies"
        Test.Use("AGENCY").Expect "AGENCIES"
        
        
        '  test "~fe"
        Test.Use("knife").Expect "knives"
        Test.Use("Knife").Expect "Knives"
        Test.Use("KNIFE").Expect "KNIVES"
        
        
        '  test ~ss
        Test.Use("glass", -1).Expect "glasses"
        Test.Use("glass", 0).Expect "No glasses"
        Test.Use("glass", 1).Expect "1 glass"
        Test.Use("glass", 8).Expect "8 glasses"
        
        
        
        '  test exceptions
        '
        Test.Use("all").Expect "all"
        Test.Use("All").Expect "All"
        Test.Use("ALL").Expect "ALL"
        
        '  test proper nouns
        '
        Test.Use("Nationwide").Expect "Nationwide"
        Test.Use("NESTLE").Expect "NESTLE"
        Test.Use("getz").Expect "getz"
End Select
End Sub



Public Function Plu_ral(singulrNoun As String _
                      , Optional quantty As Long = -1 _
                      , Optional use_No_ifZero As Boolean = True _
                      ) As String
    
    Plu_ral = Plu.ral(singulrNoun, quantty, use_No_ifZero)
    
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
    '
End Sub


Public Sub Teardown()
    '
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    'TestBed.QuitExcel
End Sub
