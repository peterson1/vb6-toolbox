VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_Replacements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseData

Private mMethodsRan&, mMethodCount&



Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)
Dim s$(), d$, v, a$
Select Case Test.MethodName
    
    Case "Join_SameResult"
        s = Rand.mWords:        d = Rand.mSyllable
        Test.Use(s, d).TestName "words, syllable"
        
        s = Rand.mSyllables:    d = Rand.mWord
        Test.Use(s, d).TestName "syllables, word"
        
        s = Rand.mPhrases:      d = Rand.mWord
        Test.Use(s, d).TestName "phrases, word"
    
    
    Case "IsNumeric_Replacement"
        v = Null: Test.Use(v).Expect(False).TestName "Null"
        v = Empty: Test.Use(v).Expect(False).TestName "Empty"
        v = 0: Test.Use(v).Expect(True).TestName v & "%"
        v = "0": Test.Use(v).Expect(True).TestName v & "$"
        v = True: Test.Use(v).Expect(False).TestName v
        v = False: Test.Use(v).Expect(False).TestName v
        
        v = 567: Test.Use(v).Expect(True).TestName v
        v = 123.45: Test.Use(v).Expect(True).TestName v
        v = "abcd": Test.Use(v).Expect(False).TestName v
        
        v = "123-": Test.Use(v).Expect(False).TestName v
        v = "-123": Test.Use(v).Expect(True).TestName v
        v = "234+": Test.Use(v).Expect(False).TestName v
        v = "+63917": Test.Use(v).Expect(False).TestName v
        v = "0917": Test.Use(v).Expect(False).TestName v
        
        v = 0.5: Test.Use(v).Expect(True).TestName v & "&"
        v = "0.5": Test.Use(v).Expect(True).TestName v & "$"
        
        v = "12e45": Test.Use(v).Expect(False).TestName v & "$ lower"
        v = "12E45": Test.Use(v).Expect(False).TestName v & "$ upper"
        v = 1.2E+46: Test.Use(v).Expect(True).TestName v & "#"
        v = "1.2E+46": Test.Use(v).Expect(True).TestName v & "$"
        v = "1e3.4": Test.Use(v).Expect(False).TestName v & "$"
        v = "1e+34": Test.Use(v).Expect(True).TestName v & "$"
        v = 1E+34: Test.Use(v).Expect(True).TestName v & "#"
        
        v = "1d34": Test.Use(v).Expect(False).TestName v & "$ lower"
        v = "1D34": Test.Use(v).Expect(False).TestName v & "$ upper"
        
        v = "1,23": Test.Use(v).Expect(False).TestName v
    
    
    Case "CustomTrim"
        Test.Use(" abc ").Expect "abc"
        Test.Use("  abc ").Expect "abc"
        Test.Use(" abc  ").Expect "abc"
        Test.Use(Space(5) & "abc" & Space(10)).Expect "abc"
        Test.Use(vbCrLf & "abc ").Expect "abc"
        Test.Use(" abc" & vbCrLf).Expect "abc"
        Test.Use("").Expect ""
        Test.Use(" ").Expect ""
        Test.Use(vbCrLf).Expect ""
    
    
    Case "GetTokens"
        d = "[a]bc": Test.Use(d).Expect(Array("a")).TestName d
        d = "a[b]c": Test.Use(d).Expect(Array("b")).TestName d
        d = "ab[c]": Test.Use(d).Expect(Array("c")).TestName d
        d = "[ab]c": Test.Use(d).Expect(Array("ab")).TestName d
        d = "a[bc]": Test.Use(d).Expect(Array("bc")).TestName d
        d = "[abc]": Test.Use(d).Expect(Array("abc")).TestName d
        d = "[a]b[c]de[f]g": Test.Use(d).Expect(Array("a", "c", "f")).TestName d
        d = "a[b]c[d][e]f[g]": Test.Use(d).Expect(Array("b", "d", "e", "g")).TestName d
        d = "a~[b]~c": Test.Use(d, "~[", "]~").Expect(Array("b")).TestName d
        d = "a~~[bcde]f": Test.Use(d, "~~[", "]").Expect(Array("bcde")).TestName d
    
    Case "Occurence"
        d = "abcde": a = "f": Test.Use(d, a).Expect(0).TestName d & " | " & a
        d = "a.b.c": a = ".": Test.Use(d, a).Expect(2).TestName d & " | " & a
        d = "a.b.c": a = "a.": Test.Use(d, a).Expect(1).TestName d & " | " & a
        d = "a.b.c": a = "c": Test.Use(d, a).Expect(1).TestName d & " | " & a
        d = "a.b.c": a = "a.b": Test.Use(d, a).Expect(1).TestName d & " | " & a
        d = "a.b.c": a = ".b.c": Test.Use(d, a).Expect(1).TestName d & " | " & a
        d = "abccc": a = "c": Test.Use(d, a).Expect(3).TestName d & " | " & a
        d = "abCcC": a = "c": Test.Use(d, a, vbTextCompare).Expect(3).TestName d & " || " & a
        d = "bbbbc": a = "b": Test.Use(d, a).Expect(4).TestName d & " | " & a
        d = "aaabc": a = "a": Test.Use(d, a).Expect(3).TestName d & " | " & a
        d = "Count<br/> occur<br/>ences of <br/>the tag <br/>.": a = "<br/>": Test.Use(d, a).Expect(4).TestName d & " | " & a
    
End Select
End Sub


Public Sub ArrayEmpty_String1D_Unredimmed()
    Dim ss$()

    Assert.That ((Not ss) = -1), Iz.True

End Sub

Public Sub ArrayEmpty_String1D_1stItemBlank()
    Dim ss$()

    ReDim ss(0)

    Assert.That ((Not ss) = -1), Iz.False

End Sub

Public Sub ArrayEmpty_String1D_2Items()
    Dim ss$()

    ReDim ss(1)
    ss(1) = "blah"

    Assert.That ((Not ss) = -1), Iz.False

End Sub

Public Sub ArrayEmpty_String1D_Erased()
    Dim ss$()

    ReDim ss(1)
    ss(1) = "blah"
    
    Erase ss
    
    Assert.That ((Not ss) = -1), Iz.True

End Sub



Public Function GetTokens(strWithTokens As String _
                        , Optional startMarkr As String = "[" _
                        , Optional endMarkr As String = "]" _
                        ) As String()
    GetTokens = Replacements.GetTokens(strWithTokens, startMarkr, endMarkr)
End Function

Public Function Occurence(strText$, findThisStr$ _
                        , Optional compareMethd As VbCompareMethod = vbBinaryCompare _
                        ) As Long
    Occurence = Replacements.Occurence(strText, findThisStr, compareMethd)
End Function


Public Function CustomTrim(strUntrimmed$) As String
    CustomTrim = Trim$(strUntrimmed)
End Function


Public Sub Join_SameResult(strArray() As String _
                         , Optional delimtr$ = vbNullString _
                         )
    Dim replacemnt$, nativeProc$
    
    ' call native method
    nativeProc = VBA.Join(strArray, delimtr)
    
    ' call replacement method
    replacemnt = Replacements.Join(strArray, delimtr)
    
    ' compare
    Assert.That replacemnt, Iz.EqualTo(nativeProc) _
             , "Result of Replacements.Join() should match native method."
End Sub


Public Function IsNumeric_Replacement(exprssion) As Boolean
    IsNumeric_Replacement = Replacements.IsNumeric(exprssion)
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
    
    mMethodsRan = mMethodsRan + 1
    If mMethodsRan = mMethodCount Then Me.FixtureTeardown
End Sub


Public Sub FixtureTeardown()
    If mMethodsRan < mMethodCount Then Exit Sub
    
End Sub
