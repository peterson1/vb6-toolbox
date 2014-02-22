VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITreeNode

Const Lf$ = vbCrLf

Private Enum TreeLeafErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As TreeLeafErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mKey        As String
Private mLabel      As String
Private mLevel      As Long
Private mInitted    As Boolean

Private mParent     As ITreeNode
Private mValues     As cCollection






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Nodes
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function ITreeNode_AddBranch(strNodeLabel As String _
                                   , branchNodeObj As ITreeNode _
                                   ) As ITreeNode
    On Error GoTo ErrH
    
    ErrRaise "Can't add branches to tree leaf."
    
ErrH: Blame "ITreeNode_AddBranch"
End Function

Private Function ITreeNode_AddLeaf(strNodeLabel As String _
                                 , leafNodeObj As ITreeNode _
                                 ) As ITreeNode
    On Error GoTo ErrH
    
    ErrRaise "Can't add leaves to tree leaf."
    
ErrH: Blame "ITreeNode_AddLeaf"
End Function

Private Property Get ITreeNode_FirstNode() As ITreeNode
    On Error GoTo ErrH
    
    ErrRaise "A TreeLeaf can't have nodes."
    
ErrH: Blame "ITreeNode_FirstNode"
End Property

Private Property Get ITreeNode_LastNode() As ITreeNode
    On Error GoTo ErrH
    
    ErrRaise "A TreeLeaf can't have nodes."
    
ErrH: Blame "ITreeNode_LastNode"
End Property

Private Property Get ITreeNode_Node(nodeIndxOrLabel As Variant _
                   ) As ITreeNode
    On Error GoTo ErrH
    
    ErrRaise "A TreeLeaf can't have nodes."
    
ErrH: Blame "ITreeNode_Node"
End Property

Private Property Get ITreeNode_HasNode(nodeIndxOrLabel As Variant) As Boolean
    ITreeNode_HasNode = False
End Property








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Level Methods
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Property Let ITreeNode_LevelOverride(RHS As Long)
    mLevel = RHS
End Property

Private Property Get ITreeNode_MaxLevel() As Long
    ITreeNode_MaxLevel = mLevel
End Property









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Values Accessor
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
            ErrorIf Not .Exists(CStr(uniqValueKey)), "No such key"
            ITreeNode_Value = .Item(CStr(uniqValueKey))
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
'    Simple Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Property Get ITreeNode_Level() As Long
    ITreeNode_Level = mLevel
End Property

Private Property Get ITreeNode_IsLeaf() As Boolean
    ITreeNode_IsLeaf = True
End Property

Private Property Let ITreeNode_Label(ByVal RHS As String)
    mLabel = RHS
End Property

Private Property Get ITreeNode_Label() As String
    ITreeNode_Label = mLabel
End Property

Private Property Get ITreeNode_Count() As Long
    ITreeNode_Count = 0
End Property

Private Property Get ITreeNode_Key() As String
    ITreeNode_Key = mKey
End Property

Private Property Get ITreeNode_Parent() As ITreeNode
    Set ITreeNode_Parent = mParent
End Property

Public Property Get TreeNode() As ITreeNode
    Set TreeNode = Me
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
    
ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Call TreeNode.Cleanup
    
ErrH: Blame "Class_Terminate"
End Sub

Private Sub ITreeNode_Cleanup()
    On Error GoTo ErrH
    
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
                   , Optional errorNumbr As TreeLeafErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As TreeLeafErrors = -1 _
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
                , Optional ByVal errorNumbr As TreeLeafErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As TreeLeafErrors = -1 _
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
