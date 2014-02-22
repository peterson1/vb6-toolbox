VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_JpegWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SUT As JpegWriter
Private mMethodsRan&, mMethodCount&



Public Sub TemporaryCover_CreatesJpeg()
    
    Dim jpgF$: jpgF = F_.TempFile(, , ".jpg")
    
    
    ' Execute method under test.
    Call SUT.TemporaryCover(Rand.mPhrase(4, 6) _
                          , Rand.mPhrase(4, 6) _
                          , jpgF _
                          , Rand.mPhrase(4))
    
    ' Verify result.
    Assert.That jpgF, Path.Exists
    
    
    'RunCommand "start " & jpgF
    
    ' Cleanup test context.
     F_.Delete jpgF
End Sub




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
    Set SUT = New JpegWriter
    Call SUT.Init
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
