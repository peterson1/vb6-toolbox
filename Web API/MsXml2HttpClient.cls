VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsXml2HttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWebApiClient

Const Lf$ = vbCrLf

Private Enum MsXml2HttpClientErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As MsXml2HttpClientErrors    '
    Description As String
End Type
Private mError As ErrorHolder


Private mUserName       As String
Private mPassword       As String

Private mContentType    As String
Private mBaseURL        As String

Private mRequest As MSXML2.XMLHTTP





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
    Dim targtUrl$, txt$, parseErr$
    On Error GoTo ErrH
    With mRequest
        targtUrl = Client.BaseURL & url_withParams_withoutBase
        
        Call .open(httpMethd, targtUrl, varAsync:=False)
        
        .setRequestHeader "Content-Type", Client.ContentType
        
        Call .send(reqstBody)
        
        txt = .ResponseText
        
        Select Case Client.StatusCode

            Case resp_OK, resp_Created
                ' ignore these

            Case Else
                If raiseErrs Then ErrRaise _
                    "Unexpected status returned." _
                        & Lf & "Status code:  [" & Client.StatusCode & "]" _
                        & Lf & "Status text:  """ & Client.StatusText & """" _
                        & Lf & "HTTP Response:  " & txt _
                        , Client.StatusCode
        End Select
        
    
        If txt = "[]" Then                                      ' return blank collection
            Set IWebApiClient_SendRequest = New_c.Collection    '  if empty response
        
        ElseIf LenB(txt) <> 0 Then
            Set IWebApiClient_SendRequest = New_c.JSONDecodeToCollection(txt, parseErr)
            ErrorIf LenB(parseErr) <> 0, "Parse error:" & Lf & parseErr
        End If
        
    End With
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
    
    Set mRequest = Nothing
    Set mRequest = New MSXML2.XMLHTTP
    
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
    IWebApiClient_ResponseText = mRequest.ResponseText
End Property

Private Property Get IWebApiClient_StatusCode() As WebApiResponseCodes
    IWebApiClient_StatusCode = mRequest.Status
End Property

Private Property Get IWebApiClient_StatusText() As String
    IWebApiClient_StatusText = mRequest.StatusText
End Property







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
    
    Set mRequest = Nothing
    
ErrH: Blame "Class_Terminate"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As MsXml2HttpClientErrors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As MsXml2HttpClientErrors = -1 _
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
                , Optional ByVal errorNumbr As MsXml2HttpClientErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As MsXml2HttpClientErrors = -1 _
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
