VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tCollection_Wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private SUT As CollectionWrapper



Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
Select Case Test.MethodName
    
    Case "BasicFunctions1"
        Test.Use "obj1"
        Test.Use "obj2", "obj2"
        Test.Use "obj1", "obj2", "obj3"
        Test.Use "obj1", "obj2", "obj3", "obj4", "obj5", "obj6", "obj7", "obj8", "obj9"
        Test.Use "obj9", "obj8", "obj7", "obj6", "obj4", "obj5", "obj3", "obj2", "obj1"
        
End Select
End Sub



Public Sub BasicFunctions1(ParamArray objNames() As Variant)
Dim obj As TestSuite
With SUT
    
    Assert.That .Count, Iz.EqualTo(0) _
              , "Should initially have no items."
    
    
    Dim i&: For i = 0 To UBound(objNames)
        
        Set obj = Nothing
        Set obj = Sim.NewTestSuite(CStr(objNames(i)))
        
        
        '  Execute method under test.
        '
        Call .Add(obj)
        
        
        Assert.That .Count, Iz.EqualTo(i + 1) _
                  , "Add() should update item Count()."
        
        Assert.That .Item(i).Name, Iz.EqualTo(objNames(i)) _
                  , "Add() should update item Count()."
    Next i
    
    Set obj = Nothing: i = 0
    
    
    
    '  Attempt enumeration.
    '
    For Each obj In SUT
        Assert.That obj.Name, Iz.EqualTo(objNames(i)) _
                 , "Collection should be enumerable."
        i = i + 1
    Next obj
    
    
    
    '  Execute method under test.
    '
    Call .Clear
    
    Assert.That .Count, Iz.EqualTo(0) _
              , "Clear() should remove all items."
    
End With
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

Public Sub Setup()
    Set SUT = New CollectionWrapper
End Sub


Public Sub Teardown()
    Set SUT = Nothing
End Sub
