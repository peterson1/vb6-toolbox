VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITreeNode

Const Lf$ = vbCrLf

Private Enum TreeBranchErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As TreeBranchErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Private mKey        As String
Private mLabel      As String
Private mLevel      As Long
Private mInitted    As Boolean

Private mParent     As ITreeNode
Private mValues     As cCollection
Private mNodes      As cCollection







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Nodes
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function ITreeNode_AddBranch(strNodeLabel As String _
                                   , branchNodeObj As ITreeNode _
                                   ) As ITreeNode
    Dim uniqNodeKey$
    On Error GoTo ErrH
    
    uniqNodeKey = BranchKey(strNodeLabel)
    
    ErrorIf mNodes.Exists(uniqNodeKey), "Node key already exists."
    Call mNodes.Add(branchNodeObj, uniqNodeKey)
    
    Call branchNodeObj.Init(uniqNodeKey, strNodeLabel, Me, mLevel + 1)
    
    Set ITreeNode_AddBranch = branchNodeObj
    
ErrH: Blame "ITreeNode_AddBranch", "Node key: """ & uniqNodeKey & """"
End Function

Private Function ITreeNode_AddLeaf(strNodeLabel As String _
                                 , leafNodeObj As ITreeNode _
                                 ) As ITreeNode
    Dim uniqNodeKey$
    On Error GoTo ErrH
    
    uniqNodeKey = LeafKey(strNodeLabel)
    
    ErrorIf mNodes.Exists(uniqNodeKey), "Node key already exists."
    Call mNodes.Add(leafNodeObj, uniqNodeKey)
    
    Call leafNodeObj.Init(uniqNodeKey, strNodeLabel, Me, mLevel + 1)
    
    Set ITreeNode_AddLeaf = leafNodeObj
    
ErrH: Blame "ITreeNode_AddLeaf"
End Function

Private Property Get BranchKey(branchLabl As String) As String
    BranchKey = ">" & branchLabl
End Property

Private Property Get LeafKey(leafLabl As String) As String
    LeafKey = "~" & leafLabl
End Property

Private Property Get ITreeNode_Node(nodeIndxOrLabel As Variant) As ITreeNode
    Dim uniqKey$
    On Error GoTo ErrH
    With mNodes
    
        If IsNumeric(nodeIndxOrLabel) Then
            ErrorIf nodeIndxOrLabel >= .Count, "Invalid index"
            Set ITreeNode_Node = .ItemByIndex(CLng(nodeIndxOrLabel))
            
            
        Else
            uniqKey = BranchKey(CStr(nodeIndxOrLabel))
            If Not .Exists(uniqKey) Then
                uniqKey = LeafKey(CStr(nodeIndxOrLabel))
                ErrorIf Not .Exists(uniqKey), "No such key"
            End If
            
            Set ITreeNode_Node = .Item(uniqKey)
        End If
    
    End With
ErrH: Blame "ITreeNode_Node"
End Property

Private Property Get ITreeNode_FirstNode() As ITreeNode
    On Error GoTo ErrH
    With mNodes
        
        ErrorIf .Count = 0, "Branch has no child nodes."
        
        Set ITreeNode_FirstNode = .ItemByIndex(0)
        
    End With
ErrH: Blame "ITreeNode_FirstNode"
End Property

Private Property Get ITreeNode_LastNode() As ITreeNode
    On Error GoTo ErrH
    With mNodes
        
        ErrorIf .Count = 0, "Branch has no child nodes."
        
        Set ITreeNode_LastNode = .ItemByIndex(.Count - 1)
        
    End With
ErrH: Blame "ITreeNode_LastNode"
End Property

Private Property Get ITreeNode_HasNode(nodeIndxOrLabel As Variant) As Boolean
    On Error GoTo ErrH
    With mNodes
        
        If IsNumeric(nodeIndxOrLabel) Then
            ITreeNode_HasNode = CLng(nodeIndxOrLabel) >= 0 _
                            And CLng(nodeIndxOrLabel) < .Count
        Else
            ITreeNode_HasNode = .Exists(BranchKey(CStr(nodeIndxOrLabel)))
            If Not ITreeNode_HasNode Then _
                ITreeNode_HasNode = .Exists(LeafKey(CStr(nodeIndxOrLabel)))
        End If
        
    End With
ErrH: Blame "ITreeNode_HasNode"
End Property








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Level Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Property Let ITreeNode_LevelOverride(RHS As Long)
    On Error GoTo ErrH
    
    mLevel = RHS
    
    Dim i&: For i = 0 To TreeNode.Count - 1
        TreeNode.Node(i).LevelOverride = mLevel + 1
    Next i
    
ErrH: Blame "ITreeNode_LevelOverride"
End Property

Private Property Get ITreeNode_MaxLevel() As Long
    On Error GoTo ErrH
    With Me.TreeNode
        
        ITreeNode_MaxLevel = .Level
        
        Dim i&: For i = 0 To .Count - 1
            
            If .Node(i).MaxLevel > ITreeNode_MaxLevel _
                Then ITreeNode_MaxLevel = .Node(i).MaxLevel
        Next i
    End With
ErrH: Blame "ITreeNode_MaxLevel"
End Property









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Values Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Property Let ITreeNode_Value(uniqValueKey As Variant _
                                   , RHS As Variant _
                                   )
    Dim uniqKey$
    On Error GoTo ErrH
    
    ErrorIf IsNumeric(uniqValueKey), "Keys cannot be numeric."
    
    uniqKey = CStr(uniqValueKey)
    
'    ErrorIf mValues.Exists(uniqKey) _
'       , "A value already uses the key:  """ & uniqValueKey & """."
    
    '  replace old value if exists
    '
    If mValues.Exists(uniqKey) Then mValues.Remove uniqKey
    
    mValues.Add RHS, uniqKey
    
ErrH: Blame "ITreeNode_Value"
End Property

Private Property Get ITreeNode_Value(uniqValueKey As Variant) As Variant
    On Error GoTo ErrH
    With mValues
    
        If IsNumeric(uniqValueKey) Then
            ErrorIf uniqValueKey >= .Count, "Invalid index"
            ITreeNode_Value = .ItemByIndex(CLng(uniqValueKey))
            
            
        Else
            'ErrorIf Not .Exists(CStr(uniqValueKey)), "No such key"
            If .Exists(CStr(uniqValueKey)) Then
                ITreeNode_Value = .Item(CStr(uniqValueKey))
            End If
        End If
    
    End With
ErrH: Blame "ITreeNode_Value", "[uniqValueKey]: """ & uniqValueKey & """"
End Property

Private Property Get ITreeNode_HasValue(valueIndxOrKey As Variant) As Boolean
    On Error GoTo ErrH
    With mValues
        
        If IsNumeric(valueIndxOrKey) Then
            ITreeNode_HasValue = CLng(valueIndxOrKey) >= 0 _
                             And CLng(valueIndxOrKey) < .Count
        Else
            ITreeNode_HasValue = .Exists(CStr(valueIndxOrKey))
        End If
        
    End With
ErrH: Blame "ITreeNode_HasValue"
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Parser
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub FromArray2D(treeArry2D As Variant)
    Dim i&, j&, itmText$, thisLevl$, nextLevl$ _
      , lBound1&, uBound1&, lBound2&, uBound2& _
      , levl0 As ITreeNode, parentLevl As ITreeNode
    On Error GoTo ErrH
    
    lBound1 = LBound(treeArry2D, 1)
    uBound1 = UBound(treeArry2D, 1)
    lBound2 = LBound(treeArry2D, 2)
    uBound2 = UBound(treeArry2D, 2)
    
    Set levl0 = Me
    
    For i = lBound1 To uBound1
        Set parentLevl = levl0
        itmText = treeArry2D(i, lBound2)
        
        For j = lBound2 + 1 To uBound2
            
            thisLevl = treeArry2D(i, j)
            If LenB(thisLevl) = 0 Then Exit For
            
            
            '  add child if none yet
            '
            If Not parentLevl.HasNode(thisLevl) Then _
                parentLevl.AddBranch thisLevl, New TreeBranch
            
            
            '  refer to this newly added level
            '
            Set parentLevl = parentLevl.Node(thisLevl)
            
            
            '  check next level
            '   - if blank, add item
            '
            nextLevl = vbNullString
            If j < UBound(treeArry2D, 2) Then _
                nextLevl = treeArry2D(i, j + 1)
            
            If LenB(nextLevl) = 0 Then _
                parentLevl.AddLeaf itmText, New TreeLeaf
            
        Next j
    Next i
    
ErrH: Blame "FromArray2D"
End Sub








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Simple Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Property Get ITreeNode_IsLeaf() As Boolean
    ITreeNode_IsLeaf = False
End Property

Private Property Let ITreeNode_Label(ByVal RHS As String)
    mLabel = RHS
End Property

Private Property Get ITreeNode_Label() As String
    ITreeNode_Label = mLabel
End Property

Public Property Get TreeNode() As ITreeNode
    Set TreeNode = Me
End Property

Private Property Get ITreeNode_Count() As Long
    ITreeNode_Count = mNodes.Count
End Property

Private Property Get ITreeNode_Key() As String
    ITreeNode_Key = mKey
End Property

Private Property Get ITreeNode_Level() As Long
    ITreeNode_Level = mLevel
End Property

Private Property Get ITreeNode_Parent() As ITreeNode
    Set ITreeNode_Parent = mParent
End Property








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ITreeNode_Init(uniqNodeKey As String _
                         , strNodeLabel As String _
                         , parentNodeObj As ITreeNode _
                         , treeNodeLevl As Long _
                         )
    On Error GoTo ErrH
    
    ErrorIf mInitted, "Init() called twice.", AlreadyInitted
    mInitted = True
    
    mKey = uniqNodeKey
    mLabel = strNodeLabel
    mLevel = treeNodeLevl
    
    Set mParent = parentNodeObj
    
ErrH: Blame "ITreeNode_Init"
End Sub









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo ErrH
    
    Set mValues = New_c.Collection(, TextCompare)
    Set mNodes = New_c.Collection(, TextCompare)
    
ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Call TreeNode.Cleanup
    
ErrH: Blame "Class_Terminate"
End Sub

Private Sub ITreeNode_Cleanup()
    Dim nod As ITreeNode
    On Error GoTo ErrH
    
    If Not mNodes Is Nothing Then
        For Each nod In mNodes
            Call nod.Cleanup
        Next nod
        Call mNodes.RemoveAll
    End If
    Set mNodes = Nothing
    
    If Not mValues Is Nothing Then mValues.RemoveAll
    Set mValues = Nothing
    
    Set mParent = Nothing
    
ErrH: Blame "ITreeNode_Cleanup"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As TreeBranchErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As TreeBranchErrors = -1 _
                  )
    If errCondition Then ErrRaise errorMsg, errorNumbr
End Sub

Private Sub SaveError()
    With mError
        If Err Then
            .HasError = True
            .Description = Err.Description
            .Number = Err.Number
            .Source = Err.Source
            
        Else
            .HasError = False
            .Description = vbNullString
            .Number = 0
            .Source = vbNullString
        End If
    End With
    Err.Clear
End Sub

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As TreeBranchErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As TreeBranchErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = TypeName(Me) & "." & currntProcedure & "()"

        If errorNumbr = -1 Then errorNumbr = .Number
        
        Select Case errorNumbr
            Case NotInitted
                errorDescrption = TypeName(Me) & " not initted." & vbCrLf _
                               & "Please call " & TypeName(Me) _
                               & ".Init() before " & currntProcedure & "."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        ErrRaise errorDescrption, errorNumbr, .Source
            
    End With
End Sub
