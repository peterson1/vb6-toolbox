VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Label As String

Public Property Get Key() As String
End Property

Public Property Get Parent() As ITreeNode
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Level Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Let LevelOverride(RHS As Long)
End Property

Public Property Get Level() As Long
End Property

Public Property Get MaxLevel() As Long
End Property



Public Property Get IsLeaf() As Boolean
End Property



Public Property Get HasNode(nodeIndxOrLabel As Variant) As Boolean
End Property

Public Property Get HasValue(valueIndxOrKey As Variant) As Boolean
End Property

Public Property Get Count() As Long
End Property



Public Property Get Node(nodeIndxOrLabel As Variant _
                       ) As ITreeNode
End Property

Public Property Get FirstNode() As ITreeNode
End Property

Public Property Get LastNode() As ITreeNode
End Property



Public Function AddBranch(strNodeLabel As String _
                        , branchNodeObj As ITreeNode _
                        ) As ITreeNode
End Function

Public Function AddLeaf(strNodeLabel As String _
                      , leafNodeObj As ITreeNode _
                      ) As ITreeNode
End Function


Public Property Get Value(uniqValueKey As Variant) As Variant
End Property

Public Property Let Value(uniqValueKey As Variant _
                        , RHS As Variant)
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor / Deconstructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(uniqNodeKey As String _
              , strNodeLabel As String _
              , parentNodeObj As ITreeNode _
              , treeNodeLevl As Long _
              )
End Sub

Public Sub Cleanup()
End Sub
