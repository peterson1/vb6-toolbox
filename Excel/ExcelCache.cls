VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const Lf$ = vbCrLf

Private Enum ExcelCacheErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As ExcelCacheErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mValues() As Variant _
      , mCells As cCollection _
      , mRows As cCollection _
      , mMaxColumn& _
      , mHelper As Excel12Helper _
      , mNames As cSortedDictionary

'Private mMarkedRows     As cCollection
'Private mMarkedValues   As cCollection
Private mRowObjects     As cCollection

Public CurrentRow&, CurrentColumn&



Public Function RowHasData(Optional ByVal rowNumbr As Long = -1 _
                         , Optional ByVal startCol As Long = -1 _
                         , Optional ByVal endCol As Long = -1 _
                         ) As Boolean
    On Error GoTo ErrH
    
    ' if none given, use Current Row
    If rowNumbr = -1 Then rowNumbr = Me.CurrentRow
    
    ' if none given, use first column
    If startCol = -1 Then startCol = LBound(mValues, 2)
    
    ' if none given, use last column
    If endCol = -1 Then endCol = UBound(mValues, 2)
    
    Dim j&: For j = startCol To endCol
        
        If LenB(mValues(rowNumbr, j)) <> 0 Then
            RowHasData = True
            Exit Function
        End If
        
    Next j
    
    
ErrH: Blame "RowHasData"
End Function


Public Sub WriteRowsToSheet(worksheetObj As Excel.Worksheet _
                          , Optional columnLettr$ = "A" _
                          , Optional startRowNum& = 1 _
                          )
    Dim v2D(), i&, j&, cols
    On Error GoTo Cleanup
    
    ReDim v2D(1 To mRows.Count, 1 To mMaxColumn)
    
    For i = LBound(v2D, 1) To UBound(v2D, 1)
        cols = mRows.ItemByIndex(i - 1)
        
        If Not IsNull(cols) Then
            
            '  if row has less columns,
            '   - resize to maxColumns
            '
            If UBound(cols) < mMaxColumn - 1 Then _
                ReDim Preserve cols(mMaxColumn - 1)
            
            
            For j = LBound(v2D, 2) To UBound(v2D, 2)
                
                v2D(i, j) = cols(j - 1)
            Next j
        End If
    Next i
    
    With worksheetObj.Range(columnLettr & startRowNum)
        .Resize(UBound(v2D, 1), UBound(v2D, 2)) = v2D
    End With
    
    Call AssignRangeNames(worksheetObj, startRowNum)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase v2D
  LoadError "WriteRowsToSheet" ', "details of error"
End Sub


Public Sub WriteValuesToSheet(worksheetObj As Excel.Worksheet _
                            , Optional columnLettr$ = "A" _
                            , Optional startRowNum& = 1 _
                            )
    On Error GoTo ErrH
    With worksheetObj.Range(columnLettr & startRowNum)
        
        .Resize(UBound(mValues, 1), UBound(mValues, 2)) = mValues
        
    End With
ErrH: Blame "WriteValuesToSheet"
End Sub


Private Sub AssignRangeNames(worksheetObj As Excel.Worksheet _
                           , startRowNum As Long _
                           )
    Dim i&, addres$, rngeNme$, colLtr$, rowNum&
    On Error GoTo ErrH
    
    If mNames.Count = 0 Then Exit Sub
    
    For i = 0 To mNames.Count - 1
        addres = mNames.ItemByIndex(i)
        rngeNme = mNames.KeyByIndex(i)
        
        colLtr = Split(addres, "$")(0)
        rowNum = CLng(Split(addres, "$")(1)) + startRowNum - 1
        
        worksheetObj.Names.Add Name:=rngeNme _
                             , RefersTo:=colLtr & rowNum
    Next i
ErrH: Blame "AssignRangeNames"
End Sub


'  Stores current row index in a collection of collections.
'
Public Sub MarkCurrentRow(collxnKey As String)
    Dim collxn As cCollection
    On Error GoTo ErrH
    
    If mMarkedRows.Exists(collxnKey) Then
        Set collxn = mMarkedRows(collxnKey)
    Else
        Set collxn = New_c.Collection(False, BinaryCompare, False)
        mMarkedRows.Add collxn, collxnKey
    End If
    
    collxn.Add Me.CurrentRow
    
ErrH: Blame "MarkCurrentRow"
End Sub


'Public Property Get MarkedRows(collxnKey As String _
'                             ) As cCollection
'    On Error GoTo ErrH
'    With mMarkedRows
'
'        ErrorIf Not .Exists(collxnKey) _
'             , "Marked rows collection not found: """ & collxnKey & """."
'
'        Set MarkedRows = .Item(collxnKey)
'
'    End With
'ErrH: Blame "MarkedRows"
'End Property


'Public Property Get MarkedValues(collxnKey As String _
'                             ) As cCollection
'    On Error GoTo ErrH
'    With mMarkedValues
'
'        ErrorIf Not .Exists(collxnKey) _
'             , "Marked values collection not found: """ & collxnKey & """."
'
'        Set MarkedValues = .Item(collxnKey)
'
'    End With
'ErrH: Blame "MarkedValues"
'End Property


Public Property Set RowObject(rowLabl As String _
                            , objReference As Object _
                            )
    On Error GoTo ErrH
    
    mRowObjects.Add objReference, rowLabl
    
    ErrorIf mRowObjects.Count <> mRows.Count _
         , "mRowObjects.Count(" & mRowObjects.Count & ")" _
           & " <> mRows.Count(" & mRows.Count & ")" _
    
ErrH: Blame "[Set]RowObject"
End Property

Public Property Get RowObjects() As cCollection
    Set RowObjects = mRowObjects
End Property


Public Sub AddCell(celValue As Variant _
                 , Optional nextRowAftrThis As Boolean = False _
                 )
    Dim i&, vv(), collxn As cCollection
    On Error GoTo Cleanup
    
    mCells.Add celValue
    
    If nextRowAftrThis Then
        ReDim vv(mCells.Count - 1)
        For i = 0 To UBound(vv)
            vv(i) = mCells.ItemByIndex(i)
        Next i
        
        mCells.RemoveAll
        
        Call AddRowArray(vv)
    End If
    
    
'    If LenB(collxnKey) <> 0 Then
'        If mMarkedValues.Exists(collxnKey) Then
'            Set collxn = mMarkedValues(collxnKey)
'        Else
'            Set collxn = New_c.Collection(False, BinaryCompare, False)
'            mMarkedValues.Add collxn, collxnKey
'        End If
'
'        collxn.Add celValue
'    End If
    
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase vv
  LoadError "AddCell" ', "details of error"
End Sub


Public Sub AddRow(ParamArray celValues() As Variant)
    Dim i&, vv
    On Error GoTo Cleanup
    
    If IsNull(celValues(0)) Then
        vv = Null
        
    Else
        ReDim vv(UBound(celValues))
        
        For i = 0 To UBound(vv)
            vv(i) = celValues(i)
        Next i
    End If
    
    Call AddRowArray(vv)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            If Not IsNull(vv) Then Erase vv
  LoadError "AddRow" ', "details of error"
End Sub


Private Sub AddRowArray(rwArray1D As Variant)
    On Error GoTo ErrH
    
    If IsNull(rwArray1D) Then
        mRows.Add Null
        
    Else
        mRows.Add rwArray1D
    
        If UBound(rwArray1D) + 1 > mMaxColumn _
            Then mMaxColumn = UBound(rwArray1D) + 1 ' update max columns
    End If
    
    Me.CurrentRow = mRows.Count
    
ErrH: Blame "AddRowArray"
End Sub


Public Sub ClearRows()
    On Error GoTo ErrH
    
    Call mRows.RemoveAll
    Call mRowObjects.RemoveAll
    
    Call mNames.RemoveAll
    
    Me.CurrentRow = 0
    
ErrH: Blame "ClearRows"
End Sub


Public Property Let RangeName(Optional ByVal rowNumbr As Long = -1 _
                            , Optional columnLetterOrNumbr As Variant _
                            , ByVal rngeName As String _
                            )
    Dim columnLettr$
    On Error GoTo ErrH
    
    If rowNumbr = -1 Then rowNumbr = Me.CurrentRow
    
    If IsMissing(columnLetterOrNumbr) Then
        columnLettr = mHelper.ColumnLetter(Me.CurrentColumn)
    
    ElseIf IsNumeric(columnLetterOrNumbr) Then
        columnLettr = mHelper.ColumnLetter(CLng(columnLetterOrNumbr))
    
    Else
        columnLettr = CStr(columnLetterOrNumbr)
    End If
    
    rngeName = mHelper.CompliantRangeName(rngeName)
    
    
    '  if it's there already...
    '
    If mNames.Exists(rngeName) Then
        Debug.Print "Duplicate RangeName: """ & rngeName & """ : " _
             & mNames(rngeName) & " vs " & columnLettr & "$" & rowNumbr
        mNames.Remove rngeName
    End If
    
    mNames.Add rngeName, columnLettr & "$" & rowNumbr

'Debug.Print rngeName, columnLettr & "$" & rowNumbr

ErrH: Blame "[Let]RangeName", " = " & rngeName
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Simple Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get Value(Optional ByVal rowIndx& = -1 _
                        , Optional ByVal colIndx& = -1 _
                        ) As Variant
Attribute Value.VB_UserMemId = 0
    On Error GoTo ErrH
    
    If rowIndx = -1 Then rowIndx = Me.CurrentRow
    If colIndx = -1 Then colIndx = Me.CurrentColumn
    
    Value = mValues(rowIndx, colIndx)
    
ErrH: Blame "[Get]Value"
End Property

Public Sub NextRow()
    Me.CurrentRow = Me.CurrentRow + 1
End Sub

Public Property Get EndOfRows() As Boolean
    On Error GoTo ErrH
    
    EndOfRows = Me.CurrentRow > UBound(mValues, 1)
    
ErrH: Blame "EndOfRows"
End Property

Public Property Get RowCount() As Long
    RowCount = mRows.Count
End Property

Public Property Get This() As ExcelCache
    Set This = Me
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub StoreUsedRangeValues(worksheetObj As Excel.Worksheet _
                              , Optional startCelAdrs$ = "A1" _
                              )
    On Error GoTo ErrH
    
    mValues = mHelper.GetUsedRange(worksheetObj).Formula
    
ErrH: Blame "StoreUsedRangeValues"
End Sub









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo ErrH
    
    Set mHelper = New Excel12Helper
    Set mRows = New_c.Collection()
    Set mCells = New_c.Collection()
    Set mNames = New_c.SortedDictionary(TextCompare)
    
'    Set mMarkedRows = New_c.Collection(False, TextCompare)
'    Set mMarkedValues = New_c.Collection(False, TextCompare)
    Set mRowObjects = New_c.Collection(False, BinaryCompare, False)
    
ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    Call Me.Cleanup
    
ErrH: Blame "Class_Terminate"
End Sub

Public Sub Cleanup()
    'Dim coll As cCollection
    On Error GoTo ErrH
    
    Set mHelper = Nothing
    
'    If Not mMarkedRows Is Nothing Then
'        For Each coll In mMarkedRows
'            coll.RemoveAll
'        Next coll
'        mMarkedRows.RemoveAll
'    End If
'    Set mMarkedRows = Nothing
'
'    If Not mMarkedValues Is Nothing Then
'        For Each coll In mMarkedValues
'            coll.RemoveAll
'        Next coll
'        mMarkedValues.RemoveAll
'    End If
'    Set mMarkedValues = Nothing
    
    If Not mRowObjects Is Nothing Then mRowObjects.RemoveAll
    Set mRowObjects = Nothing

    If Not mNames Is Nothing Then mNames.RemoveAll
    Set mNames = Nothing
    
    If Not mCells Is Nothing Then mCells.RemoveAll
    Set mCells = Nothing
    
    If Not mRows Is Nothing Then mRows.RemoveAll
    Set mRows = Nothing
    
    Erase mValues
    
ErrH: Blame "Cleanup"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As ExcelCacheErrors = -1 _
                  )
    If errCondition Then Err.Raise errorNumbr, TypeName(Me), errorMsg
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
                , Optional ByVal errorNumbr As ExcelCacheErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As ExcelCacheErrors = -1 _
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

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub

