VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_XmlWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseData

Private SUT As XmlWriter _
      , mDOM As cSimpleDOM
Private mMethodsRan&, mMethodCount&


Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
Dim n&, u$, e$, a$()
Select Case Test.MethodName
    
    Case "LetA_StoresMultipleAttributes"
        n = 1: Test.Use(n).TestName Plu.ral("attribute", n)
        n = 3: Test.Use(n).TestName Plu.ral("attribute", n)
        n = 5: Test.Use(n).TestName Plu.ral("attribute", n)
        n = 6: Test.Use(n).TestName Plu.ral("attribute", n)
        n = 20: Test.Use(n).TestName Plu.ral("attribute", n)
        
    
    Case "Encode_Decode"
        u = "less < than & greater > than"
        e = "less &lt; than &amp; greater &gt; than"
        Test.Use(u, e).TestName u & " --> " & e
        
        u = "a single-quoted 'buzzword'"
        e = "a single-quoted &apos;buzzword&apos;"
        Test.Use(u, e).TestName u & " --> " & e
    
        u = "a double-quoted """ & "buzzword" & """"
        e = "a double-quoted &quot;buzzword&quot;"
        Test.Use(u, e).TestName u & " --> " & e
        
        u = "St. Josemaría Escrivá"
        e = "St. Josemar&iacute;a Escriv&aacute;"
        Test.Use(u, e).TestName u & " --> " & e
        
    
    Case "XmlSafe"

        u = " ": Test.Use(u).Expect(True).TestName "space """ & u & """"        ' 32
        u = "!": Test.Use(u).Expect(True).TestName "exclamation """ & u & """"  ' 33
        
        u = """": Test.Use(u).Expect(False).TestName "double-quote " & u        ' 34
        
        u = "#": Test.Use(u).Expect(True).TestName "hash """ & u & """"         ' 35
        u = "$": Test.Use(u).Expect(True).TestName """" & u & """"              ' 36
        u = "%": Test.Use(u).Expect(True).TestName """" & u & """"              ' 37
        
        u = "&": Test.Use(u).Expect(False).TestName """" & u & """"             ' 38
        u = "'": Test.Use(u).Expect(False).TestName """" & u & """"             ' 39
        
        u = "(": Test.Use(u).Expect(True).TestName """" & u & """"              ' 40
        u = ")": Test.Use(u).Expect(True).TestName """" & u & """"              ' 41
        u = "*": Test.Use(u).Expect(True).TestName """" & u & """"              ' 42
        u = "+": Test.Use(u).Expect(True).TestName """" & u & """"              ' 43
        u = ",": Test.Use(u).Expect(True).TestName """" & u & """"              ' 44
        u = "-": Test.Use(u).Expect(True).TestName """" & u & """"              ' 45
        u = ".": Test.Use(u).Expect(True).TestName """" & u & """"              ' 46
        u = "/": Test.Use(u).Expect(True).TestName """" & u & """"              ' 47
        u = "0": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 48
        u = "1": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 49
        u = "2": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 50
        u = "3": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 51
        u = "4": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 52
        u = "5": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 53
        u = "6": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 54
        u = "7": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 55
        u = "8": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 56
        u = "9": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 57
        u = ":": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 58
        u = ";": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 59
        
        u = "<": Test.Use(u).Expect(False).TestName "number """ & u & """"      ' 60
        
        u = "=": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 61
        
        u = ">": Test.Use(u).Expect(False).TestName "number """ & u & """"      ' 62
        
        u = "?": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 63
        u = "@": Test.Use(u).Expect(True).TestName "number """ & u & """"       ' 64
        u = "A": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 65
        u = "B": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 66
        u = "C": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "D": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "E": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "F": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "G": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "H": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "I": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "J": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "K": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "L": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "M": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "N": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "O": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "P": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "Q": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "R": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "S": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "T": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "U": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "V": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "W": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "X": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"
        u = "Y": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 89
        u = "Z": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 90
        u = "[": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 91
        u = "\": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 92
        u = "]": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 93
        u = "^": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 94
        u = "_": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 95
        u = "`": Test.Use(u).Expect(True).TestName "uppercase """ & u & """"    ' 96
        u = "a": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 97
        u = "b": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 98
        u = "c": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "d": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "e": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "f": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "g": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "h": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "i": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "j": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "k": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "l": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "m": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "n": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "o": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "p": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "q": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "r": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "s": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "t": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "u": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "v": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "w": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "x": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"
        u = "y": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 121
        u = "z": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 122
        u = "{": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 123
        u = "|": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 124
        u = "}": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 125
        u = "~": Test.Use(u).Expect(True).TestName "lowercase """ & u & """"    ' 126
        
    
    
    Case "GetTag_CreatesNewTag"
        Test.Use("oneTag").TestName "1 tag"
        Test.Use("tag1", "tag2").TestName "2 different tags"
        Test.Use("tagN", "tagN").TestName "2 identical tags"
        Test.Use("t1", "t2", "t3", "t4", "t5", "t6", "t7", "t8").TestName "8 tags"
        Test.Use("t1", "t2", "t3", "t4", "t5", "t6", "t7", "t8", "t9").TestName "9 tags"
        Test.Use("t", "t", "t", "t", "t", "t", "t", "t", "t", "t", "t" _
               , "t", "t", "t", "t", "t", "t", "t", "t", "t", "t", "t").TestName "22 tags"
    
    
    Case "SetTag_CreatesNewTag"
        Test.Use("oneTag").TestName "1 tag"
        Test.Use("tag1", "tag2").TestName "2 different tags"
        Test.Use("tagN", "tagN").TestName "2 identical tags"
        Test.Use("t1", "t2", "t3", "t4", "t5", "t6", "t7", "t8").TestName "8 tags"
        Test.Use("t1", "t2", "t3", "t4", "t5", "t6", "t7", "t8", "t9").TestName "9 tags"
        Test.Use("t", "t", "t", "t", "t", "t", "t", "t", "t", "t", "t" _
               , "t", "t", "t", "t", "t", "t", "t", "t", "t", "t", "t").TestName "22 tags"
    
        
    Case "LetAttributes_SetsAttributes"
        Test.Use("""", "att", "val").TestName "double-quote, 1 attr"
        Test.Use("""", "k1", "v1", "k2", "v2").TestName "double-quote, 2 atts"
        
        
        
    Case "SplitPairs"
        ReDim a(2, 1):  a(0, 0) = "k1":     a(0, 1) = "v01"
                        a(1, 0) = "k02":    a(1, 1) = "v002"
                        a(2, 0) = "k003":   a(2, 1) = "v3"
        
        u = "k1='v01' k02='v002' k003='v3'"
        Test.Use(u, "=", "'", " ").Expect(a).TestName u
        
        u = "k1:'v01' k02:'v002' k003:'v3'"
        Test.Use(u, ":", "'", " ").Expect(a).TestName u
        
        u = "k1:v01 k02:v002 k003:v3"
        Test.Use(u, ":", "", " ").Expect(a).TestName u
        
        'u = "k1: v01 k02: v002 k003: v3"
        'Test.Use(u, ":", "", " ").Expect(a).TestName u
        
        u = "k1 = 'v01' k02= 'v002' k003 ='v3'"
        Test.Use(u, "=", "'", " ").Expect(a).TestName u
        
End Select
End Sub


Public Function SplitPairs(joindPairs$ _
                         , keyValSeparatr$ _
                         , valQuotes$ _
                         , pairDelimitr$ _
                         ) As String()
    SplitPairs = SUT.SplitPairs(joindPairs, keyValSeparatr _
                              , valQuotes, pairDelimitr)
End Function

Public Sub LetAttributes_SetsAttributes(strQuote As String _
                                      , ParamArray keyVals() As Variant _
                                      )
    Dim i&, joind$, ss$(), attCount&
    
    attCount = (UBound(keyVals) + 1) / 2
    
    ReDim ss(attCount - 1)
    For i = 0 To UBound(keyVals) Step 2
        ss(i \ 2) = keyVals(i) & "=" _
                  & strQuote & keyVals(i + 1) & strQuote
    Next i
    joind = Join(ss, " ")
    
    
    ' call method
    Let SUT.Attributes = joind
    
    
    Assert.That SUT.AttributeCount, Iz.EqualTo(attCount) _
             , "Should update attributes count."
    
    
    For i = 0 To UBound(keyVals) Step 2
        Assert.That SUT.AttributeExists(keyVals(i)), Iz.True _
                 , "Should add each attribute."
    Next i
End Sub


Public Sub GetLevel_ReturnsLevel()
    Dim obj As XmlWriter
    
    '  verify root level
    '
    Assert.That SUT.Level, Iz.EqualTo(0) _
             , "Root level should be 0."
    
    
    '  add a tag to root
    '
    Set obj = SUT.Tag("level_1_a")
    Assert.That obj.Level, Iz.EqualTo(1) _
             , "1st tag added to root should be level 1."
    
    
    '  add another tag to root
    '
    Set obj = SUT.Tag("level_1_b")
    Assert.That obj.Level, Iz.EqualTo(1) _
             , "2nd tag added to root should be level 1."
    
    
    '  add a tag to level 1 tag
    '
    Set obj = obj.Tag("level_2")
    Assert.That obj.Level, Iz.EqualTo(2) _
             , "Tag added to level 1 should be level 2."
    
    
    ' Cleanup test context.
    Set obj = Nothing
End Sub



Public Function XmlSafe(charactr$) As Boolean
    XmlSafe = SUT.XmlSafe(charactr)
End Function



Public Sub Encode_Decode(origCharactr As String _
                       , expectdText As String)

    Assert.That SUT.Encode(origCharactr) _
              , Iz.EqualTo(expectdText) _
             , "Encode() should replace XML-encodeable characters."

    Assert.That SUT.Decode(expectdText) _
              , Iz.EqualTo(origCharactr) _
             , "Decode() should parse XML-encoded characters."
End Sub


Public Sub LetTagName_SetsTagName()
    Dim actualVal$, expctd$
    
    expctd = Rand.mWord()
    
    
    ' Execute method under test.
    SUT.Name = expctd
    
    
    ' Verify result.
    actualVal = Root.TagName
    Assert.That actualVal, Iz.EqualTo(expctd)
    
End Sub

Public Sub LetTagName2x_ChangesTagName()
    Dim actualVal$, expctd$
    
    '  Init once
    SUT.Name = "firstRoot"
    
    
    expctd = "secondRoot"
    
    
    ' Init twice
    SUT.Name = expctd
    
    
    ' Verify result.
    actualVal = Root.TagName
    Assert.That actualVal, Iz.EqualTo(expctd)
    
End Sub



Public Sub LetA_StoresMultipleAttributes(attrbuteCount As Long)
    
    '  generate random values for each attribute
    '
    Dim attKeys$(): ReDim attKeys(attrbuteCount - 1)
    Dim attVals$(): ReDim attVals(attrbuteCount - 1)
    Dim i&: For i = 0 To UBound(attKeys)
        attKeys(i) = "att" & i + 1
        attVals(i) = Rand.mWord
    Next i
    
    
    '  verify emptiness
    Assert.That SUT.AttributeCount, Iz.EqualTo(0)
    
    
    '  store all
    For i = 0 To UBound(attKeys)
        SUT.A_(attKeys(i)) = attVals(i)
        
        Assert.That SUT.AttributeCount, Iz.EqualTo(i + 1) _
                 , "AttributeCount should reflect added attribute."
        
        Assert.That SUT.AttributeExists(attKeys(i)), Iz.True
    Next i
    
    
    '  parse all
    '
    Dim parsd() As cAttribute:  ReDim parsd(UBound(attKeys))
    With Root.Attributes
        Assert.That .Count, Iz.EqualTo(UBound(attKeys) + 1) _
                  , "Tag should have " & UBound(attKeys) + 1 & " attributes."
        
        For i = 0 To UBound(parsd)
            Set parsd(i) = .Item(i)
        Next i
    End With
    
    
    ' verify all
    For i = 0 To UBound(parsd)
        
        Assert.That parsd(i).Name, Iz.EqualTo(attKeys(i)) _
                , "[Let]A_ should store attribute keys."
    
        Assert.That parsd(i).Text, Iz.EqualTo(attVals(i)) _
                , "[Let]A_ should store attribute values."
    
        Assert.That SUT.A_(attKeys(i)), Iz.EqualTo(attVals(i)) _
                , "[Get]A_ should return attribute value."
    Next i
End Sub


Public Sub LetA_ChangesExistingAttribute()
    
    Dim attKey$:    attKey = "attributeKey"
    Dim oldVal$:    oldVal = "old value"
    Dim newVal$:    newVal = "new value"
    
    ' store old value
    SUT.A_(attKey) = oldVal
    Assert.That SUT.A_(attKey), Iz.EqualTo(oldVal)
    
    ' store new value with same key
    SUT.A_(attKey) = newVal
    
    ' Verify result.
    Assert.That Root.Attributes.Count, Iz.EqualTo(1) _
             , "Should not create new attribute."
             
    Assert.That SUT.A_(attKey), Iz.EqualTo(newVal) _
            , "[Let]A_ should change value for existing attribute."
End Sub



' restricts names starting with numbers
' escapes encodable entities

Public Sub GetTag_CreatesNewTag(ParamArray tagNmes() As Variant)
    Dim i&, tagObj As XmlWriter
    
    
    ' Execute method under test.
    For i = 0 To UBound(tagNmes)
        Set tagObj = SUT.Tag(CStr(tagNmes(i)))
        
        Assert.That tagObj, Iz.[Not].Nothing _
                 , "[Get]Tag should return something."
    Next i
    
    
    ' Verify result.
    With Root.ChildElements
        
        Assert.That .Count, Iz.EqualTo(UBound(tagNmes) + 1) _
                 , "[Get]Tag should add new tag with each call."
        
        For i = 0 To UBound(tagNmes)
            Assert.That .Item(i).TagName, Iz.EqualTo(tagNmes(i)) _
                     , "[Get]Tag should set name of each new tag."
        Next i
    End With
End Sub

Public Sub SetTag_CreatesNewTag(ParamArray tagNmes() As Variant)
    Dim i&, tagObj As XmlWriter
    Const s$ = "Text for "
    
    For i = 0 To UBound(tagNmes)
        Set tagObj = Nothing
        Set tagObj = New XmlWriter
        tagObj.InnerText = s & tagNmes(i)
        
        ' Execute method under test.
        Set SUT.Tag(CStr(tagNmes(i))) = tagObj
    Next i
    
    
    ' Verify result.
    With Root.ChildElements
        
        Assert.That .Count, Iz.EqualTo(UBound(tagNmes) + 1) _
                 , "[Set]Tag should add new tag with each call."
        
        For i = 0 To UBound(tagNmes)
            Assert.That .Item(i).TagName, Iz.EqualTo(tagNmes(i)) _
                     , "[Set]Tag should set name of each new tag."
            
            Assert.That .Item(i).Text, Iz.EqualTo(s & tagNmes(i)) _
                     , "[Set]Tag should set the text of each new tag."
        Next i
    End With
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Test Utility Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function Root() As cElement
    
    Set mDOM = New_RC4.SimpleDOM(SUT.ToString)
    
    Set Root = mDOM.Root
    
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
    
    Set SUT = New XmlWriter
    SUT.Name = "root"
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


