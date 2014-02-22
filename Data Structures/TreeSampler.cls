VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeSampler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const Lf$ = vbCrLf

Private Enum SampleTreesErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As SampleTreesErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Public Enum SampleTrees
    tr_CreatureNames
End Enum



Public Function AsTreeNodes(samplTree As SampleTrees _
                          ) As ITreeNode
    Dim treeBrnch As TreeBranch
    On Error GoTo ErrH
    
    'Set AsTreeNodes = ToTreeNodes(AsArray(samplTree))
    Set treeBrnch = New TreeBranch
    Call treeBrnch.FromArray2D(Me.AsArray(samplTree))
    
    Set AsTreeNodes = treeBrnch
    
ErrH: Blame "AsTreeNodes"
End Function


'Public Function ToTreeNodes(treeArry2D As Variant _
'                          ) As ITreeNode
'    Dim i&, j&, itmText$, thisLevl$, nextLevl$ _
'      , lBound1&, uBound1&, lBound2&, uBound2& _
'      , levl0 As ITreeNode, parentLevl As ITreeNode
'    On Error GoTo ErrH
'
'    lBound1 = LBound(treeArry2D, 1)
'    uBound1 = UBound(treeArry2D, 1)
'    lBound2 = LBound(treeArry2D, 2)
'    uBound2 = UBound(treeArry2D, 2)
'    Set levl0 = New TreeBranch
'
'    For i = lBound1 To uBound1
'        Set parentLevl = levl0
'        itmText = treeArry2D(i, lBound2)
'
'        For j = lBound2 + 1 To uBound2
''If i = 254 And j = 3 Then
''    Debug.Print "f"
''End If
'
'            thisLevl = treeArry2D(i, j)
'            If LenB(thisLevl) = 0 Then Exit For
'
'
'            '  add child if none yet
'            '
'            If Not parentLevl.HasNode(thisLevl) Then _
'                parentLevl.AddBranch thisLevl, New TreeBranch
'
'
'            '  refer to this newly added level
'            '
'            Set parentLevl = parentLevl.Node(thisLevl)
'
'
'            '  check next level
'            '   - if blank, add item
'            '
'            nextLevl = vbNullString
'            If j < UBound(treeArry2D, 2) Then _
'                nextLevl = treeArry2D(i, j + 1)
'
'            If LenB(nextLevl) = 0 Then
'
'                'If parentLevl.HasNode(itmText) Then
'                    'Debug.Print "meron na: " & itmText
'                'Else
'                    parentLevl.AddLeaf itmText, New TreeLeaf
'                'End If
'            End If
'
'        Next j
'    Next i
'
'    Set ToTreeNodes = levl0
'
'ErrH: Blame "ToTreeNodes", "i:" & i & "  j:" & j
'End Function


Private Function ToCollection(treeArry2D() As String _
                            ) As cCollection
    Dim i&, j&, itmText$ _
      , thisLevl$, nextLevl$ _
      , levl0 As cCollection, parentLevl As cCollection
    On Error GoTo ErrH
    
    Set levl0 = New_c.Collection(, TextCompare)
    
    For i = 0 To UBound(treeArry2D, 1)
        Set parentLevl = levl0
        itmText = treeArry2D(i, 0)
        
        For j = 1 To UBound(treeArry2D, 2)
            
            thisLevl = treeArry2D(i, j)
            If LenB(thisLevl) = 0 Then Exit For
            
            
            '  add child if none yet
            '
            If Not parentLevl.Exists(thisLevl) Then _
                parentLevl.Add New_c.Collection(, TextCompare), thisLevl
            
            
            '  refer to this newly added level
            '
            Set parentLevl = parentLevl(thisLevl)
            
            
            '  check next level
            '   - if blank, add item
            '
            nextLevl = vbNullString
            If j < UBound(treeArry2D, 2) Then _
                nextLevl = treeArry2D(i, j + 1)
            
            If LenB(nextLevl) = 0 Then _
                parentLevl.Add itmText
            
        Next j
    Next i
    
    Set ToCollection = levl0
    
ErrH: Blame "ToCollection"
End Function


Public Function AsArray(samplTree As SampleTrees _
                      ) As String()
    On Error GoTo ErrH
    Select Case samplTree

        Case tr_CreatureNames:      AsArray = CreatureNames()

    End Select
ErrH: Blame "AsArray"
End Function





Private Function CreatureNames() As String()
    Dim c As New cCollection
    On Error GoTo Cleanup

'   Insect
'       Garapata
'           "Willie"
'           "Vice"
'
'   Reptile
'       Dinosaur
'           T-Rex
'               "Muncher"
'           Stegosaur
'               "Denver"
'           Bronto
'               "Little Foot"
'               "Little Foot's mom"
'       Crocodile
'           "Dundee"
'       Snake
'           "Slithery"
'
'   Bird
'       "Big Bird"
'       Eagle
'           "My Eagle"
'       Pugo
'           "His Pugo"
'       Itik
'           "Her Itik 1"
'           "Her Itik 2"
'           "Her Itik 3"
'
'   Mammal
'       Human
'           The Simpsons
'               "Homer"
'               "Marge"
'               "Bart"
'               "Lisa"
'               "Maggie"
'           "Leng-Leng"
'           "Paw-Paw
'       Dog
'           "Maguyguy"
'           "Manyunyu"
'           "Winchee"
'
    c.Add " Willie          :Insect     :Garapata"
    c.Add " Vice            :Insect     :Garapata"
    
    c.Add " Muncher         :Reptile    :Dinosaur   :T-Rex"
    c.Add " Denver          :Reptile    :Dinosaur   :Stegosaur"
    c.Add " Lil Foot        :Reptile    :Dinosaur   :Bronto"
    c.Add " Lil Foot's mom  :Reptile    :Dinosaur   :Bronto"
    c.Add " Dundee          :Reptile    :Crocodile"
    c.Add " Slithery        :Reptile    :Snake"
    
    c.Add " Big Bird        :Bird"
    c.Add " My Eagle        :Bird       :Eagle"
    c.Add " His Pugo        :Bird       :Pugo"
    c.Add " Her Itik 1      :Bird       :Itik"
    c.Add " Her Itik 2      :Bird       :Itik"
    c.Add " Her Itik 3      :Bird       :Itik"
    
    c.Add " Homer           :Mammal     :Human      :The Simpsons"
    c.Add " Marge           :Mammal     :Human      :The Simpsons"
    c.Add " Bart            :Mammal     :Human      :The Simpsons"
    c.Add " Lisa            :Mammal     :Human      :The Simpsons"
    c.Add " Maggie          :Mammal     :Human      :The Simpsons"
    c.Add " Leng-Leng       :Mammal     :Human"
    c.Add " Paw-Paw         :Mammal     :Human"
    c.Add " Maguyguy        :Mammal     :Dog"
    c.Add " Manyunyu        :Mammal     :Dog"
    c.Add " Winchee         :Mammal     :Dog"
    
    CreatureNames = ArrayCodeTo2D(c)
    
Cleanup:    SaveError
            Set c = Nothing
            'Erase someArray
  LoadError "CreatureNames" ', "details of error"
End Function




Private Function ArrayCodeTo2D(arrayCodeCollxn As cCollection _
                             , Optional delimitr As String = ":" _
                             ) As String()
    Dim i&, j&, ss$(), t2D$(), maxColBound&
    On Error GoTo Cleanup
    With arrayCodeCollxn
        
        
        '  initial loop gets max columns
        '
        For i = 0 To .Count - 1
            ss = Split(.ItemByIndex(i), delimitr)
            
            If UBound(ss) > maxColBound Then _
                maxColBound = UBound(ss)
        Next i
        
        
        ReDim t2D(.Count - 1, maxColBound)
        
        
        '  2nd loop fills the 2D array
        '
        For i = 0 To UBound(t2D, 1)
            ss = Split(.ItemByIndex(i), delimitr)
            
            For j = 0 To UBound(ss)
                t2D(i, j) = Trim$(ss(j))
            Next j
        Next i
        
        
        ArrayCodeTo2D = t2D
    
    End With
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss, t2D
  LoadError "ArrayCodeTo2D" ', "details of error"
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    'On Error GoTo ErrH
    
    'Set mSomeObject = New Something
    
'ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Call Me.Cleanup
    
ErrH: Blame "Class_Terminate"
End Sub

Public Sub Cleanup()
    On Error GoTo ErrH
    
    'Set mSomeObject = Nothing
    
ErrH: Blame "Cleanup"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As SampleTreesErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As SampleTreesErrors = -1 _
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
                , Optional ByVal errorNumbr As SampleTreesErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As SampleTreesErrors = -1 _
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

