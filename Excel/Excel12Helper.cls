VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Excel12Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum Excel12HelperErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As Excel12HelperErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Const clr_White& = 16777215     'RGB(255, 255, 255)
Const clr_RedFill& = 13158655   'RGB(255, 200, 200)
Const clr_RedLine& = 6579455    'RGB(255, 100, 100)
Const clr_OrangeFill& = 9948159 'RGB(255, 203, 151)
Const clr_OrangeLine& = 29670   'RGB(230, 115,   0)
Const clr_GreenFill& = 10741703 'RGB(199, 231, 163)
Const clr_GreenLine& = 4640397  'RGB(141, 206,  70)
Const clr_Grey5& = 15921906     'RGB(242, 242, 242)
Const clr_Grey15& = 14211288    'RGB(216, 216, 216)
Const clr_Grey25& = 12566463    'RGB(191, 191, 191)
Const clr_Grey35& = 10855845    'RGB(165, 165, 165)


Private Enum ValidNameCharacters
    upr_A = 65:     low_a = 97:     asc_0 = 48:     undrscr = 95
    upr_Z = 90:     low_z = 122:    asc_9 = 57:     asc_dot = 46
End Enum


Public Enum ColorThemes
    thmeBlank
    thmeNormal
    thmeRed
    thmeOrange
    thmeGreen
    thmeGrey1
    thmeGrey2
    thmeGrey3
End Enum



Public Property Let ApplyTheme(celRange As Excel.Range _
                             , colrTheme As ColorThemes _
                             )
    On Error GoTo ErrH
    Dim fillColr&, lineColr&
    With celRange
    
    If colrTheme = thmeBlank Then
        .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).ThemeColor = 1
        .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).ThemeColor = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).ThemeColor = 1
        .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).ThemeColor = 1
        .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous:  .Borders(xlInsideHorizontal).ThemeColor = 1
    
    ElseIf colrTheme = thmeNormal Then
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        '.Borders(xlInsideVertical).LineStyle = xlNone
    
    Else
        Select Case colrTheme
        
            Case thmeRed
                fillColr = clr_RedFill
                lineColr = clr_RedLine
                
            Case thmeOrange
                fillColr = clr_OrangeFill
                lineColr = clr_OrangeLine
    
            Case thmeGreen
                fillColr = clr_GreenFill
                lineColr = clr_GreenLine
    
            Case thmeGrey1
                fillColr = clr_Grey5
                lineColr = clr_Grey15
    
            Case thmeGrey2
                fillColr = clr_Grey15
                lineColr = clr_Grey25
    
            Case thmeGrey3
                '.Font.Color = clr_White
                fillColr = clr_Grey25
                lineColr = clr_Grey35
        
        End Select
        .Borders(xlEdgeTop).Color = lineColr
        .Borders(xlEdgeRight).Color = lineColr
        .Borders(xlEdgeBottom).Color = lineColr
        .Borders(xlEdgeLeft).Color = lineColr
        .Borders(xlInsideHorizontal).Color = lineColr
        
        .Borders(xlInsideVertical).Color = fillColr
        .Interior.Color = fillColr
    End If
    
    End With
ErrH: Blame "ShowDefaultBorder"
End Property


Public Property Get FindSheet(workbookObj As Excel.Workbook _
                            , sheetNme As String _
                            ) As Excel.Worksheet
    On Error GoTo ErrH
    
    ErrorIf workbookObj.Sheets.Count = 0, "Workbook has no sheets."
    
    For Each FindSheet In workbookObj.Sheets
        
        If UCase$(Trim$(FindSheet.Name)) _
         = UCase$(Trim$(sheetNme)) Then Exit Property
    
    Next FindSheet
    
    Set FindSheet = Nothing
    Err.Raise 9999, , "Worksheet not found:  """ & sheetNme & """."
    
ErrH: Blame "FindSheet", "filename:  " & workbookObj.Name
End Property


Public Property Get GetUsedRange(worksheetObj As Excel.Worksheet _
                               ) As Excel.Range
    Dim lastRowNum&, lastColNum&
    On Error GoTo ErrH
    
    Const startCelAdrs$ = "A1" 'won't work if starting from any other cell
    
    With worksheetObj
        
        lastRowNum = .Cells.Find(What:="*" _
                              , After:=.Range(startCelAdrs) _
                             , LookIn:=xlFormulas _
                             , LookAt:=xlWhole _
                        , SearchOrder:=xlByRows _
                    , SearchDirection:=xlPrevious _
                    ).Row
        
        lastColNum = .Cells.Find(What:="*" _
                              , After:=.Range(startCelAdrs) _
                             , LookIn:=xlFormulas _
                             , LookAt:=xlWhole _
                        , SearchOrder:=xlByColumns _
                    , SearchDirection:=xlPrevious _
                    ).Column
        
        Set GetUsedRange = .Range(startCelAdrs, _
                                 .Cells(lastRowNum, lastColNum))
    End With
ErrH: Blame "GetUsedRange"
End Property


Public Function AddNewSheetDeleteOld(workbookObj As Excel.Workbook _
                                   ) As Excel.Worksheet
    Dim sht As Excel.Worksheet
    On Error GoTo ErrH
    
    Set AddNewSheetDeleteOld = workbookObj.Worksheets.Add()
    
    For Each sht In workbookObj.Worksheets
        
        If sht.Name <> AddNewSheetDeleteOld.Name Then sht.Delete
        
    Next sht
    
ErrH: Blame "AddNewSheetDeleteOld"
End Function



Public Property Get ColumnNumber(columnLettr As String) As Long
    On Error GoTo ErrH
    
    ErrorIf Len(columnLettr) > 1 _
         , "Only single letters are currently supported."
    
    'Asc("A") = 65
    ColumnNumber = (Asc(UCase$(columnLettr)) - 65) + 1
    
ErrH: Blame "[Get]ColumnNumber"
End Property


Public Function CopySheet(sourceWrkbook As Workbook _
                        , sourceSheetNme As String _
                        ) As Excel.Worksheet
    Dim srcSheet As Worksheet
    On Error GoTo Cleanup
    
    With sourceWrkbook
        Set srcSheet = .Worksheets(sourceSheetNme)
        Dim i&: i = srcSheet.Index
        
        srcSheet.Copy srcSheet
        
        Set CopySheet = .Worksheets(i)
        
    End With
Cleanup:    SaveError
            Set srcSheet = Nothing
            'Erase someArray
  LoadError "CopySheet" ', "details of error"
End Function


Public Property Get NamedRange(wrkSheet As Excel.Worksheet _
                             , rangeNme As String _
                             ) As Excel.Range
    Dim nme As Excel.Name
    On Error GoTo ErrH
    
    With wrkSheet
'        For Each nme In .Names
'            If UCase$(nme.Name) = UCase$("'" & wrkSheet.Name & "'!" & rangeNme) Then
'
'                Set NamedRange = .Range(Evaluate(nme.RefersTo))
'                Exit Property
'
'            End If
'        Next nme
        
        Set NamedRange = .Range(Evaluate(.Names(rangeNme).RefersTo))
        
    End With
    
    
    'ErrRaise "No such Named Range: """ & rangeNme & """."
    
ErrH:
    If Err.Number = 1004 Then
        Err.Clear
        ErrRaise "No such Named Range: """ & rangeNme & """."
    End If
    
    Blame "NamedRange"
End Property

'Public Property Get NamedRange(wrkSheet As Excel.Worksheet _
'                             , rangeNme As String _
'                             ) As Excel.Range
'    Dim nme As Excel.Name
'    On Error GoTo ErrH
'
'    With wrkSheet
'        For Each nme In .Names
'            If UCase$(nme.Name) = UCase$("'" & wrkSheet.Name & "'!" & rangeNme) Then
'
'                Set NamedRange = .Range(Evaluate(nme.RefersTo))
'                Exit Property
'
'            End If
'        Next nme
'    End With
'
'    ErrRaise "No such Named Range: """ & rangeNme & """."
'
'ErrH: Blame "NamedRange"
'End Property



Public Property Get CompliantRangeName(ByVal strText As String _
                                     , Optional defaultReplacementChar$ = "_" _
                                     ) As String
    Dim i&, c$
    On Error GoTo ErrH
    
    For i = 1 To Len(strText)
        c = Mid$(strText, i, 1)
        
        Select Case Asc(c)
            
            Case upr_A To upr_Z     '
            Case low_a To low_z     '  do nothing,
            Case asc_0 To asc_9     '   - these are all valid characters
            Case undrscr, asc_dot   '
            
            
            '  if none of the above,
            '   - swap with replacement char
            '
            Case Else:  strText = Replace(strText, c, defaultReplacementChar)
        End Select
        
    Next i
    
    
    '  fix doubled underscores
    '
    If InStr(strText, "__") Then strText = Replace(strText, "__", "_")
    
    CompliantRangeName = strText
    
ErrH: Blame "CompliantRangeName"
End Property



Public Function DrawTreeNodes(treeNodesObj As ITreeNode _
                            , worksheetObj As Excel.Worksheet _
                            , startColumnLettr As String _
                            , startRowNumbr As Long _
                            , Optional aplyCosmetics As Boolean = True _
                            ) As ExcelCache
    Dim i&, xlCache As ExcelCache
    On Error GoTo ErrH
    
    Set xlCache = New ExcelCache
    
    TreeNodesToCache treeNodesObj, xlCache, 0
    
    xlCache.WriteRowsToSheet worksheetObj, startColumnLettr, startRowNumbr
    
    If aplyCosmetics Then
        Me.FormatEntireSheet worksheetObj
        
        FormatTreeColumns worksheetObj, startColumnLettr _
                                     , treeNodesObj.MaxLevel
        
        FormatTreeRows xlCache, worksheetObj, startColumnLettr _
                                             , startRowNumbr _
                                             , treeNodesObj.MaxLevel
    End If
    
    Set DrawTreeNodes = xlCache
    
ErrH: Blame "DrawTreeNodes"
End Function


Private Sub TreeNodesToCache(treeNodesObj As ITreeNode _
                           , xlCache As ExcelCache _
                           , currentIndx As Long _
                           )
    Dim i&, s$
    On Error GoTo ErrH
    With treeNodesObj
    
        For i = 2 To .Level     '  add blank cells
            xlCache.AddCell Null        '   to indent to node level
        Next i                          '
        
        '  conclude row with actual string
        xlCache.AddCell .Label, True
        Set xlCache.RowObject(.Label) = treeNodesObj
        
        
        
        For i = 0 To .Count - 1
            TreeNodesToCache .Node(i), xlCache, i '  recurse thru children
        Next i
        
        
        '  if total row needs to be added...
        '
        If currentIndx = .Parent.Count - 1 And .Level > 0 Then
            
            
            '  add spacer rows above [Total row]
            AddSpacerRow xlCache, 4 - .Level
            
            
            '  add [Total row]
            '
            For i = 2 To .Level - 1
                xlCache.AddCell Null
            Next i
            s = "Total " & .Parent.Label
            xlCache.AddCell s, True
            Set xlCache.RowObject(s) = Nothing
            
            
            
            '  add spacer rows below [Total row]
            AddSpacerRow xlCache, 5 - .Level
            
        End If
        
    End With
ErrH: Blame "TreeNodesToCache"
End Sub


Private Sub AddSpacerRow(xlCacheObj As ExcelCache _
                       , Optional lRowCount As Long = 1 _
                       )
    Dim i&
    On Error GoTo ErrH
    With xlCacheObj
        For i = 1 To lRowCount
            
            Call .AddRow(Null)
            
            Set .RowObject("") = Nothing
            
        Next i
    End With
ErrH: Blame "AddSpacerRow"
End Sub


Private Sub FormatTreeRows(xlCacheObj As ExcelCache _
                         , worksheetObj As Excel.Worksheet _
                         , startColumnLettr As String _
                         , startRowNumbr As Long _
                         , maxLevl As Long _
                         )
    Dim i&, labl$, obj As ITreeNode, rowObjs As cCollection
    On Error GoTo ErrH
    
    Set rowObjs = xlCacheObj.RowObjects
    For i = 0 To rowObjs.Count - 1
        labl = rowObjs.KeyByIndex(i)
        Set obj = rowObjs.ItemByIndex(i)
        
        With worksheetObj.Range(startColumnLettr & startRowNumbr + i)
            
            
            '  if it's a normal row
            '
            If Not obj Is Nothing Then
                .RowHeight = 30 - obj.Level * 3.5
                
                If obj.IsLeaf Then
                    
                    If obj.Level <> maxLevl Then
                        With .Offset(, obj.Level - 1)
                            Me.ApplyTheme(.Resize(, (maxLevl - obj.Level) + 1)) = thmeNormal
                        End With
                    Else
                        If IsEmpty(obj.Parent.Value("group address")) Then
                            
                            With .Offset(, obj.Level - 1).Resize(obj.Parent.Count)
                                Me.ApplyTheme(.Cells) = thmeNormal
                                obj.Parent.Value("group address") = StripColumnLetters(.Cells)
                            End With
                        End If
                    End If
                End If
                
            
            '  if it's a Totals row...
            '
            ElseIf LenB(labl) <> 0 And obj Is Nothing Then
                With .Resize(, maxLevl)
                    .IndentLevel = 7
                    .RowHeight = 25
                    .VerticalAlignment = xlBottom
                End With
            End If
            
        End With
    Next i
ErrH: Blame "FormatTreeRows"
End Sub


Public Function StripColumnLetters(cellRnge As Range _
                                 , Optional replacemnt As String = "[c]" _
                                 ) As String
    On Error GoTo ErrH
    
    With cellRnge.Rows
        StripColumnLetters = replacemnt _
                           & .Item(1).Row _
                           & ":" _
                           & replacemnt _
                           & .Item(.Rows.Count).Row
    End With
ErrH: Blame "StripColumnLetters"
End Function


Private Sub FormatTreeColumns(worksheetObj As Excel.Worksheet _
                            , startColumnLettr As String _
                            , maxLevl As Long _
                            )
    On Error GoTo ErrH
    With worksheetObj.Range(startColumnLettr & ":" & startColumnLettr)
        
        
        '  format all except last column
        '
        With .Resize(, maxLevl - 1)
            .ColumnWidth = 3
        End With
        
        
        '  format last column
        '
        With .Offset(, maxLevl - 1)
            .EntireColumn.AutoFit
            If .ColumnWidth < 15 Then .ColumnWidth = 15
            If .ColumnWidth > 50 Then .ColumnWidth = 50
        End With
        
        
        '  format level 0
        '
        With .Offset(, 0)
            .Font.Name = "Cambria"
            .Font.Size = 13
            .Font.Bold = True
        End With
        
        '  format level 1
        '
        With .Offset(, 1)
            .Font.Name = "Cambria"
            .Font.Size = 13
            .Font.Italic = True
        End With
        
        '  format level 2
        '
        With .Offset(, 2)
            .Font.Name = "Cambria"
            .Font.Size = 12
        End With
        
        '  format level 3
        '
        With .Offset(, 3)
            .Font.Size = 9
            .Font.Bold = True
        End With
        
        '  format level 4
        '
        With .Offset(, 4)
            .Font.Size = 9
            .Font.Italic = True
        End With
        
        '  format level 5
        '
        With .Offset(, 5)
            .Font.Size = 8
        End With
        
    End With
ErrH: Blame "FormatTreeColumns"
End Sub


Public Sub FormatEntireSheet(worksheetObj As Excel.Worksheet)
    On Error GoTo ErrH
    With worksheetObj.Cells
    
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        
        Me.ApplyTheme(.Cells) = thmeBlank
        
    End With
ErrH: Blame "FormatEntireSheet"
End Sub


Public Property Get ColumnLetter(columnIndx As Long _
                               , Optional ascStart As Long = 64 _
                               , Optional maxIndx As Long = 26 _
                               ) As String
    On Error GoTo ErrH
    
    If columnIndx > maxIndx Then
        If (columnIndx Mod maxIndx) = 0 Then
            ColumnLetter = ColumnLetter((columnIndx / maxIndx) - 1) & ColumnLetter(maxIndx)
        Else
            ColumnLetter = ColumnLetter(columnIndx \ maxIndx) & ColumnLetter(columnIndx Mod maxIndx)
        End If
    Else
        ColumnLetter = Chr(ascStart + columnIndx)
    End If
    
ErrH: Blame "[Get]ColumnLetter"
End Property


Public Sub FreezePane(worksheetObj As Excel.Worksheet _
                    , cellAddrss As String _
                    )
    On Error GoTo ErrH
    
    With worksheetObj.Application
        
        .ScreenUpdating = True
        
        worksheetObj.Activate
        
        worksheetObj.Range(cellAddrss).Select
        
        .ActiveWindow.FreezePanes = True
        
    End With
    
ErrH: Blame "FreezePane"
End Sub


' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init()

End Sub









' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Class Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub Class_Initialize()
    On Error GoTo ErrH
    
    'Set mSomeObject = New Something
    
ErrH: Blame "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    
    'Set mApp = Nothing
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As Excel12HelperErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As Excel12HelperErrors = -1 _
                  )
    'If errCondition Then Err.Raise errorNumbr, TypeName(Me), errorMsg
    If errCondition Then ErrRaise errorMsg, errorNumbr, TypeName(Me)
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
                , Optional ByVal errorNumbr As Excel12HelperErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As Excel12HelperErrors = -1 _
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

