VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_cElementWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseData

Private SUT As cElementWrapper _
      , mDOM As cSimpleDOM

Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
Select Case Test.MethodName

    Case "Find"
        Test.Use("id=item4").Expect("Text of item4").TestName "id: no space between"
        Test.Use("id =item3").Expect("Text of item3").TestName "id: space before"
        Test.Use("id= item2").Expect("Text of item2").TestName "id: space after"
        Test.Use("id = item1").Expect("Text of item1").TestName "id: space both"
        
        Test.Use("att=att4").Expect("Text of item4").TestName "att: no space between"
        Test.Use("att =att3").Expect("Text of item3").TestName "att: space before"
        Test.Use("att= att2").Expect("Text of item2").TestName "att: space after"
        Test.Use("att = att1").Expect("Text of item1").TestName "att: space both"
    
    Case "Tag"
        Test.Use("a").Expect "item a"
        Test.Use("c").Expect "item c"
        Test.Use("1").Expect "item b"
        Test.Use("3").Expect "item d"
        Test.Use(0).Expect "item a"
        Test.Use(2).Expect "item c"
    
    Case "Text"
        Test.Use("<a>123</a>").Expect "123"
        Test.Use("<b><a>123</a></b>").Expect "<a>123</a>"
        Test.Use("<b><a1>123</a1><a2>234</a2></b>").Expect "<a1>123</a1><a2>234</a2>"
        Test.Use("<br />").Expect ""
        Test.Use("<br/>").Expect ""
    
    Case "Enumerable"
        Test.Use "oneTag"
        Test.Use "tag1", "tag2"
        Test.Use "t1", "t2", "t3", "t4", "t5", "t6", "t7"
    
    
    Case "ChildCount"
        Test.Use("<r><a>Hi!</a></r>").Expect 1
        Test.Use("<r><br/></r>").Expect 1
        Test.Use("<r><br/><br/></r>").Expect 2
        Test.Use("<r><a1></a1><a2></a2></r>").Expect 2
        Test.Use("<r></r>").Expect 0
        Test.Use("<selfClose />").Expect 0
    
End Select
End Sub


Public Function ChildCount(xmlStr$) As Long
    
    Call SUT.Init(Parse(xmlStr))
    
    ChildCount = SUT.ChildCount
    
End Function

'Public Sub Enumerable(ParamArray tagNmes() As Variant)
'    Dim i&, e As Object
'
'
'    '  compose and load XML
'    '
'    With New_Xml("root")
'        For i = 0 To UBound(tagNmes)
'            .Tag(tagNmes(i)).InnerText = Rand.mWord
'        Next i
'
'        Call SUT.Init(Parse(.ToString))
'    End With
'
'
'    '  attempt Enumeration
'    '
'    For Each e In SUT
'
'        Assert.That e.Name, Iz.EqualTo(tagNmes(i)) _
'                 , "Should be support 'For-Each' enumeration."
'
'        i = i + 1
'    Next e
'
'    Set e = Nothing
'End Sub



Public Function Text(xmlStr As String) As String
    
    Call SUT.Init(Parse(xmlStr))
    
    Text = SUT.Text
End Function

Public Function Tag(tagNameOrIndx) As String
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "    <a>item a</a>" _
        & "    <b>item b</b>" _
        & "    <c>item c</c>" _
        & "    <d>item d</d>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Tag = SUT.Tag(tagNameOrIndx).Text
    
End Function


Public Function Find(searchFiltr$) As String
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "    <i id='item1' att='att1'>Text of item1</i>" _
        & "    <i id='item2' att='att2'>Text of item2</i>" _
        & "    <i id='item3' att='att3'>Text of item3</i>" _
        & "    <i id='item4' att='att4'>Text of item4</i>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Find = SUT.Find(searchFiltr).Text
    
End Function


Public Sub Find_ErrorIfNotFound()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "    <i id='item1' att='att1'>Text of item1</i>" _
        & "    <i id='item2' att='att2'>Text of item2</i>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.Throws cElementWrapperErrors.ItemNotFound
    
    Call SUT.Find("id = item3")
    
End Sub

Public Sub Find_NoErrorIfNotFound()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "    <i id='item1' att='att1'>Text of item1</i>" _
        & "    <i id='item2' att='att2'>Text of item2</i>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.That SUT.Find("id = item3", False), Iz.Nothing _
             , "Should not raise error if we don't want it to."
End Sub

Public Sub Find_ErrorIfNoChildren()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.Throws cElementWrapperErrors.NoChildren
    
    Call SUT.Find("child=any")
    
End Sub

Public Sub Find_NoErrorIfNoChildren()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.That SUT.Find("child=any", False), Iz.Nothing _
             , "Should not raise error if we don't want it to."
End Sub




Public Sub Tag_ErrorIfNotFound()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "    <tag1>Text of tag1</tag1>" _
        & "    <tag2>Text of tag2</tag2>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.Throws cElementWrapperErrors.ItemNotFound
    
    Call SUT.Tag("tag3")
    
End Sub

Public Sub Tag_NoErrorIfNotFound()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "    <tag1>Text of tag1</tag1>" _
        & "    <tag2>Text of tag2</tag2>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.That SUT.Tag("tag3", False), Iz.Nothing _
             , "Should not raise error if we don't want it to."
End Sub

Public Sub Tag_ErrorIfNoChildren()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.Throws cElementWrapperErrors.NoChildren
    
    Call SUT.Tag("child")
    
End Sub

Public Sub Tag_NoErrorIfNoChildren()
    Dim xmlStr$: xmlStr _
        = "<root>" _
        & "</root>"
    
    Call SUT.Init(Parse(xmlStr))
    
    Assert.That SUT.Tag("child", False), Iz.Nothing _
             , "Should not raise error if we don't want it to."
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Test Utility Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function Parse(xmlStr$) As cElement
    
    Set mDOM = Nothing
    Set mDOM = New_RC4.SimpleDOM(xmlStr)
    
    Set Parse = mDOM.Root
    
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
    Set SUT = New cElementWrapper
    
End Sub


Public Sub Teardown()
    Set SUT = Nothing
    Set mDOM = Nothing
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
    'TestBed.QuitExcel
End Sub
