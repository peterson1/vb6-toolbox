VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbRcHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWebApiClient

Const Lf$ = vbCrLf

Private Enum vbRcHttpClientErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As vbRcHttpClientErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Private WithEvents mDownloads As cDownloads


Private mUserName       As String
Private mPassword       As String

Private mContentType    As String
Private mBaseURL        As String

Private mResponseText   As String
Private mStatusText     As String
Private mStatusCode     As WebApiResponseCodes




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Main Method
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function IWebApiClient_SendRequest(httpMethd As String _
                                         , url_withParams_withoutBase As String _
                                         , Optional reqstBody As String _
                                         , Optional raiseErrs As Boolean = True _
                                         ) As vbRichClient5.cCollection
    Dim targtUrl$, parseErr$, methd As DownloadMethod
    On Error GoTo ErrH
    With mDownloads
        targtUrl = Client.BaseURL & url_withParams_withoutBase
        
        Select Case Trim$(UCase$(httpMethd))
            Case "GET":     methd = dl_HTTP_GET
            Case "POST":    methd = dl_HTTP_POST
            Case "PUT":     methd = dl_HTTP_PUT
            Case "DELETE":  methd = dl_HTTP_DELETE
        End Select
        
        
        .Download(targtUrl, methd, reqstBody _
               , "Content-Type: " & Client.ContentType _
               ).WaitForCompletion 3
    End With
    
'    mStatusCode = 0
'    Do Until mStatusCode <> 0
'        DoEvents
'    Loop
    
    
    Select Case mStatusCode

        Case resp_OK, resp_Created
            ' ignore these

        Case Else
            If raiseErrs Then ErrRaise _
                "Unexpected status returned." _
                    & Lf & "Status code:  [" & Client.StatusCode & "]" _
                    & Lf & "Status text:  """ & Client.StatusText & """" _
                    & Lf & "HTTP Response:  " & mResponseText _
                    , Client.StatusCode
    End Select
    

    If mResponseText = "[]" Then                                      ' return blank collection
        Set IWebApiClient_SendRequest = New_c.Collection    '  if empty response
    
    ElseIf LenB(mResponseText) <> 0 Then
        Set IWebApiClient_SendRequest = New_c.JSONDecodeToCollection(mResponseText, parseErr)
        ErrorIf LenB(parseErr) <> 0, "Parse error:" & Lf & parseErr
    End If
        
ErrH: Blame "IWebApiClient_SendRequest"
End Function






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub IWebApiClient_Init(wbApiBaseUrl As String _
                             , Optional userNme As String _
                             , Optional passwrd As String _
                             , Optional contentTyp As String = "application/json; charset=utf-8" _
                             )
    On Error GoTo ErrH
    
    Set mDownloads = Nothing
    Set mDownloads = New_c.Downloads
    
    Client.BaseURL = wbApiBaseUrl
    Client.UserName = userNme
    Client.Password = passwrd
    Client.ContentType = contentTyp
    
ErrH: Blame "IWebApiClient_Init"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Simple Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get Client() As IWebApiClient
    Set Client = Me
End Property

Private Property Let IWebApiClient_BaseURL(ByVal RHS As String)
    mBaseURL = RHS
End Property

Private Property Get IWebApiClient_BaseURL() As String
    IWebApiClient_BaseURL = mBaseURL
End Property

Private Property Let IWebApiClient_ContentType(ByVal RHS As String)
    mContentType = RHS
End Property

Private Property Get IWebApiClient_ContentType() As String
    IWebApiClient_ContentType = mContentType
End Property

Private Property Let IWebApiClient_UserName(ByVal RHS As String)
    mUserName = RHS
End Property

Private Property Get IWebApiClient_UserName() As String
    IWebApiClient_UserName = mUserName
End Property

Private Property Let IWebApiClient_Password(ByVal RHS As String)
    mPassword = RHS
End Property

Private Property Get IWebApiClient_Password() As String
    IWebApiClient_Password = mPassword
End Property

Private Property Get IWebApiClient_ResponseText() As String
    IWebApiClient_ResponseText = mResponseText
End Property

Private Property Get IWebApiClient_StatusCode() As WebApiResponseCodes
    IWebApiClient_StatusCode = mStatusCode
End Property

Private Property Get IWebApiClient_StatusText() As String
    IWebApiClient_StatusText = mStatusText
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Download Events
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub mDownloads_DownloadComplete(downloadObj As vbRichClient5.cDownload _
                                      , ByVal errNum As Long _
                                      , errString As String)
    On Error GoTo ErrH
    With downloadObj
        mStatusCode = CLng(.QueryInfo(qiStatusCode))
        mStatusText = .QueryInfo(qiStatusText)
        mResponseText = StrConv(.GetContentData, vbUnicode)
    End With
ErrH: Blame "mDownloads_DownloadComplete"
End Sub







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
    
    If Not mDownloads Is Nothing Then
        With mDownloads
            If .Count <> 0 Then .AbortAllDownloads
            Call .removeAll
        End With
    End If
    
    Set mDownloads = Nothing
    
ErrH: Blame "Cleanup"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As vbRcHttpClientErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As vbRcHttpClientErrors = -1 _
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
                , Optional ByVal errorNumbr As vbRcHttpClientErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As vbRcHttpClientErrors = -1 _
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
