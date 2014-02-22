VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JpegWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum JpegWriterErrors           ' you may make this Public for tests
    ErrorBase = vbObjectError + 513     ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ExpressionTooComplex = 16
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As JpegWriterErrors    '
    Description As String
End Type
Private mError As ErrorHolder



Public Sub TemporaryCover(bookTtle As String _
                        , authrName As String _
                        , outputJpgPath As String _
                        , Optional subTtle$ = "(subtitle goes here)" _
                        , Optional jpgWidth As Long = 600 _
                        , Optional jpgHeight As Long = 800 _
                        )
    Dim srf As cCairoSurface, c As cCairoContext _
      , boxW#, boxH#, innrX#, innrY#, innrW#, innrH#, nextY#, div2#
    On Error GoTo Cleanup
    
    With New_RC4.Cairo
        Set srf = .CreateSurface(jpgWidth, jpgHeight)
    End With
    Set c = srf.CreateContext
    With c
        
        
        '  draw border
        '
        .SetSourceColor vbWhite ' bg color
        .Paint
        
        Const boxX# = 60
        Const boxY# = 60
        boxW = jpgWidth - boxX * 2
        boxH = jpgHeight - boxY * 2
        .Rectangle boxX, boxY, boxW, boxH
        .SetSourceColor vbBlack
        .Stroke
        
        innrX = boxX + 40
        innrY = boxY + 70
        innrW = jpgWidth - innrX * 2
        innrH = jpgHeight - innrY * 2
        '.Rectangle innrX, innrY, innrW, innrH: .Stroke ' uncomment this to show inner rectangle
        
        
        '  draw title text
        '
        .SelectFont "Times New Roman", 40
        nextY = DrawText(bookTtle, c, innrX, innrY, innrW)
        
        
        '  draw subtitle text
        '
        .SelectFont "Times New Roman", 20
        nextY = DrawText(subTtle, c, innrX, nextY + 50, innrW)
        
        
        '  divide remaining space among them
        '
        div2 = (innrH - nextY) / 2
        
        
        '  draw "temporary" text
        '
        .SelectFont "verdana", 14, vbRed
        DrawText "(temporary book cover)", c, innrX, nextY + div2, innrW
        
        
        '  draw author text
        '
        .SelectFont "Times New Roman", 20
        DrawText "by:  " & authrName, c, innrX, nextY + div2 * 2, innrW
        
    End With
    
    Call srf.WriteContentToJpgFile(outputJpgPath)
    
Cleanup:    SaveError
            Set c = Nothing
            Set srf = Nothing
            'Erase someArray
  LoadError "TemporaryCover" ', "details of error"
End Sub



Private Function DrawText(strText As String _
                        , cairoContxt As cCairoContext _
                        , xPosition As Double _
                        , yPosition As Double _
                        , maxmumWidth As Double _
                        , Optional textAlignmnt As AlignmentConstants = vbCenter _
                        ) As Double
    Dim numRows&, extntH#, redrawAttmpts&
    On Error GoTo Cleanup
    With cairoContxt
        
        
        '  calculate height of text-extent
        '
        numRows = .DrawText(0, 0, maxmumWidth, 0, strText, , textAlignmnt)
        extntH = numRows * .GetFontHeight()
        
        
        '  render actual text
        '
RedrawText:
        .DrawText xPosition, yPosition, maxmumWidth, extntH, strText, , textAlignmnt
        
        
        '  return next clear yPosition
        '
        DrawText = yPosition + extntH
        
    End With
Cleanup:    If Err = ExpressionTooComplex And _
               redrawAttmpts < 1 Then
                redrawAttmpts = redrawAttmpts + 1
                Err.Clear
                GoTo RedrawText
            End If
            SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "DrawText" ', "details of error"
End Function



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
    On Error GoTo Cleanup
    
    'Set mSomeObject = New Something
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo Cleanup
    
    'Set mSomeObject = Nothing
    
Cleanup:    SaveError
            'Set someObj = Nothing
            'Erase someArray
  LoadError "Class_Terminate"
End Sub





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As JpegWriterErrors = -1 _
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

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As JpegWriterErrors = -1 _
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

