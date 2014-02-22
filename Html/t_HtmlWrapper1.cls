VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_HtmlWrapper1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseData

Private SUT As HtmlWrapper1
Private mTempFile$
Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
Select Case Test.MethodName
    
    Case "MethodName_GoesHere"
    
End Select
End Sub


Public Sub BodyTags_Enumerable()
    Dim i&, tNames$(), tTexts$(), e As cElementWrapper
    
    i = Rand.mNumber(1, 10)
    tNames = Rand.mWords(i, i)
    tTexts = Rand.mWords(i, i)
    
    
    '  compose test html
    '
    With New_Xml("html")
        .Tag("head").InnerText = "empty"
        With .Tag("body")
            
            For i = 0 To UBound(tNames)
                .Tag(tNames(i)).InnerText = tTexts(i)
            Next i
            
        End With
Debug.Print .ToString(True)
        Call .ToFile(mTempFile, True)
    End With
    
    ' parse
    Call SUT.Init(mTempFile)
    
    ' verify
    i = 0
    For Each e In SUT.BodyTags()
        
        Assert.That e.Name, Iz.EqualTo(tNames(i))
        Assert.That e.Text, Iz.EqualTo(tTexts(i))
        
        i = i + 1
    Next e
    
    Erase tNames, tTexts
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
    
    mTempFile = F_.TempFile("test", "temp", ".html")
    
    Set SUT = New HtmlWrapper1
    'Call SUT.Init( mPath$
End Sub


Public Sub Teardown()
    Set SUT = Nothing
    
    F_.Delete mTempFile
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    'TestBed.QuitExcel
End Sub
