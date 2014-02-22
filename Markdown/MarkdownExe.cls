VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarkdownExe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const Lf$ = vbCrLf
Const Lf2$ = vbCrLf & vbCrLf

Private Enum MarkdownExeErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Public Enum MarkdownParsers
    parsr_MultiMarkdown_1
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As MarkdownExeErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mExePath$ _
      , mParser As MarkdownParsers


Public Function ToHtml(markDownTxt As String) As String
    On Error GoTo ErrH
    Select Case mParser

        Case parsr_MultiMarkdown_1
            ToHtml = MmdToHtml_v1(markDownTxt, True)
            
    End Select
ErrH: Blame "ToHtml"
End Function


Private Function MmdToHtml_v1(markDownTxt As String _
                            , Optional processContentsOfRawHtml As Boolean = False _
                            , Optional generateIdAttribsForHeadrs As Boolean = True _
                            ) As String
    Dim tmpF$, cmd$
    On Error GoTo ErrH
    
    tmpF = F_.MakeTempFile(EscapeEntities(markDownTxt))
    
    cmd = """" & mExePath & """"
    
    If processContentsOfRawHtml Then cmd = cmd & " --process-html"
    
    If Not generateIdAttribsForHeadrs Then cmd = cmd & " --nolabels"
    
    cmd = cmd & " """ & tmpF & """"
    
    MmdToHtml_v1 = PostProcess_v1(GetCommandOutput(cmd))
    
    F_.Delete tmpF
    
ErrH: Blame "MmdToHtml_v1"
End Function


Private Function PostProcess_v1(markdownProcessedStr As String _
                              ) As String
    Dim dom As cSimpleDOM, i&, ss$(), s$
    On Error GoTo Cleanup
    
    If LenB(Trim(markdownProcessedStr)) = 0 Then Exit Function
    
    Set dom = New_RC4.SimpleDOM("<body>" & markdownProcessedStr & "</body>")
    With dom.Root.ChildElements
        ReDim ss(.Count - 1)
        
        For i = 0 To UBound(ss)
            With E_(.Item(i))
                
                If .Name = "blockquote" Then
                    ss(i) = ReplaceBlockQuoteTag_v1(.Self)
                
                ElseIf .Name = "pre" Then
                    ss(i) = ReplacePreTag_v1(.Self)
                
                ElseIf .A_("class", "") = "footnotes" Then
                    ss(i) = TweakFootnotes_v1(.Self)
                
                Else
                    ss(i) = .Element.XML
                End If
                
                
                '  remove square brackets from footnote links
                '
                s = " class=""footnote"">["
                If InText(ss(i), s) Then
                    ss(i) = Replace(ss(i), s, " class=""footnote"">")
                    ss(i) = Replace(ss(i), "]</a>", " </a>")
                End If
                
            End With
        Next i
    End With
    
    PostProcess_v1 = Join(ss, vbCrLf & vbCrLf)
    
Cleanup:    SaveError
            Set dom = Nothing
            Erase ss
  LoadError "PostProcess_v1" ', "details of error"
End Function


Private Function ReplaceBlockQuoteTag_v1(blockQuotTag As cElementWrapper _
                                       ) As String
    On Error GoTo ErrH
    
    ReplaceBlockQuoteTag_v1 = "<hr class='top'/>" _
                       & Lf & "<blockquote><blockquote>" _
                            & Lf & blockQuotTag.Text _
                       & Lf & "</blockquote></blockquote>" _
                       & Lf & "<hr class='bottom'/>"
    
ErrH: Blame "ReplaceBlockQuoteTag_v1"
End Function

Private Function ReplacePreTag_v1(preTagElm As cElementWrapper _
                                ) As String
    On Error GoTo ErrH
    
    ReplacePreTag_v1 = "<div class='indented'>" & Lf _
                            & Replace(preTagElm.Tag("code").Text _
                             , "  " & Lf, "<br>" & Lf) _
                     & Lf & "</div>"
    
ErrH: Blame "ReplacePreTag_v1"
End Function


Private Function TweakFootnotes_v1(footNotesDivTagElm As cElementWrapper _
                                 ) As String
    Dim i&, ss$(), xmlW As XmlWriter, li As cElement
    On Error GoTo Cleanup
    
    Set xmlW = New_Xml("div")
    xmlW.A_("class") = "footnotes"
    
    With xmlW.Tag("h3")
        .A_("id") = "endnote-title"
        .InnerText = "Endnotes"
    End With
    
    With xmlW.Tag("ol")
        For Each li In footNotesDivTagElm.Tag("ol").Children
            
            With .Tag("li")
                
                .A_("class") = "endnote"
                
                .A_("id") = E_(li).A_("id")
                
                '  replace the back arrow symbol
                '   (not always renderable) with a simpler text
                '
                .InnerText = Replace(E_(li).Tag("p").Text _
                                  , "&#160;&#8617;", "[back]")
            End With
            
        Next li
    End With
    
    TweakFootnotes_v1 = xmlW.ToString(True)
    
Cleanup:    SaveError
            Set xmlW = Nothing
            Set li = Nothing
            Erase ss
  LoadError "TweakFootnotes_v1" ', "details of error"
End Function


Private Function EscapeEntities(strText As String) As String
    Dim i&, c$, ss$()
    On Error GoTo Cleanup
    
    If LenB(strText) = 0 Then Exit Function
    
    ReDim ss(Len(strText) - 1)
    
    For i = 0 To UBound(ss)
        c = Mid$(strText, i + 1, 1)
        
        If XmlSafe(c) Then
            ss(i) = c
        Else
            ss(i) = EntityEncode(c)
        End If
    Next i
    
    EscapeEntities = Join(ss)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "EscapeEntities" ', "details of error"
End Function


Private Function XmlSafe(charactr$) As Boolean
    Dim a%: a = Asc(charactr)
    
    XmlSafe = True
    
    If a > 31 And a < 127 Then Exit Function     ' ? @ A-Z [ \ ] ^ _ ` a-z { | } ~

    'If a = 32 Then Exit Function                ' space
    If a = 9 Then Exit Function                 ' tab
    If a = 10 Then Exit Function                ' line feed
    If a = 13 Then Exit Function                ' carriage return
    
    'If a = 60 Then Exit Function                ' less-than sign
    'If a = 62 Then Exit Function                ' greater-than sign
    
    'If a > 31 And a < 63 Then Exit Function     ' ' ( ) * + , - . /  0-9 : ; =
    'If a > 32 And a < 38 Then Exit Function     ' ! " # $ %
    'If a = 61 Then Exit Function                ' =
    
    
    XmlSafe = False
End Function

Private Function EntityEncode(singleCharactr$) As String
    Dim a%
    On Error GoTo ErrH
    
    ' shortcut
    EntityEncode = "&#" & Asc(singleCharactr) & ";"
    Exit Function
    
    a = Asc(singleCharactr)
    
    Select Case a

'        Case 9:     EntityEncode = vbTab   ' tab
'        Case 10:    EntityEncode = vbLf    ' line feed
'        Case 13:    EntityEncode = vbCr    ' carriage return
'        Case 34:    EntityEncode = "&quot;"     ' double-quote
'        Case 38:    EntityEncode = "&amp;"      ' ampersand
'        Case 39:    EntityEncode = "&apos;"     ' apostrophe
'        Case 60:    EntityEncode = "&lt;"       ' less-than sign
'        Case 62:    EntityEncode = "&gt;"       ' greater-than sign
        Case 133:   EntityEncode = "&hellip;"   ' horizontal ellipsis (= three dot leader)
        Case 145:   EntityEncode = "&lsquo;"    ' left single quote ‘
        Case 146:   EntityEncode = "&rsquo;"    ' right single quote ’
        Case 147:   EntityEncode = "&ldquo;"    ' left double quote “
        Case 148:   EntityEncode = "&rdquo;"    ' right double quote ”
        Case 150:   EntityEncode = "&ndash;"    ' en dash –
        Case 151:   EntityEncode = "&mdash;"    ' em dash —
        Case 225:   EntityEncode = "&aacute;"   ' á : Latin small letter a with acute accent
        Case 237:   EntityEncode = "&iacute;"   ' í : Latin small letter i with acute accent
        Case 243:   EntityEncode = "&oacute;"   ' ó : Latin small letter o with acute accent
        
                
        Case Else
            Err.Raise -1, , "No entity encoding for character """ _
                           & singleCharactr & """ : Asc(" & a & ")."
    End Select
    
ErrH: Blame "EntityEncode"
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(appExePath As String _
              , parserTypAndVersion As MarkdownParsers)
    On Error GoTo ErrH
    
    ErrorIf Not F_.Found(appExePath) _
             , "No such file:" & vbCrLf & appExePath
    
    mExePath = appExePath
    
    mParser = parserTypAndVersion
    
ErrH: Blame "Init"
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
    
    'Set mEncoder = Nothing
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As MarkdownExeErrors = -1 _
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
                , Optional ByVal errorNumbr As MarkdownExeErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As MarkdownExeErrors = -1 _
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

