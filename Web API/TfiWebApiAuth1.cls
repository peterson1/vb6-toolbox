VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TfiWebApiAuth1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWebApiClient

Const Lf$ = vbCrLf
Const HEADER_USERNAME$ = "TFI-API-Auth-Username"
Const HEADER_AUTHSCHEME$ = "TFI-API-Auth-v1"

Private Enum TfiWebApiAuth1Errors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As TfiWebApiAuth1Errors    '
    Description As String
End Type
Private mError As ErrorHolder


Private mUserName       As String
Private mPassword       As String
Private mPasswordHash   As String

Private mContentType    As String
Private mBaseURL        As String

Private mRequest        As MSXML2.XMLHTTP
Private mCrypt          As vbRichClient5.cCrypt





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
    Dim txt$, parseErr$, reqDate As Date, contntMD5$, signture$
    On Error GoTo ErrH
    With mRequest
        
        reqDate = GmtDate(Now)
        contntMD5 = GetContentMD5(reqstBody)
        signture = GetRequestSignature(httpMethd, contntMD5, reqDate _
                                     , url_withParams_withoutBase)
        
        Call .open(httpMethd, Client.BaseURL & url_withParams_withoutBase _
                                                        , varAsync:=False)
        .setRequestHeader "Date", GmtFormat(reqDate)
        .setRequestHeader "Content-Type", Client.ContentType
        .setRequestHeader HEADER_USERNAME, Client.UserName
        .setRequestHeader "Authorization", HEADER_AUTHSCHEME & " " & signture
        If LenB(contntMD5) <> 0 Then _
            .setRequestHeader "Content-MD5", contntMD5
        
        
        Call .send(reqstBody)
        
        
        txt = .ResponseText
        
        Select Case Client.StatusCode

            Case resp_OK, resp_Created
                ' ignore these
            
            Case resp_Unauthorized
                Debug.Print .getResponseHeader("debug")
            
            Case Else
                If raiseErrs Then ErrRaise _
                    "Unexpected status returned." _
                        & Lf & "Status code:  [" & Client.StatusCode & "]" _
                        & Lf & "Status text:  """ & Client.StatusText & """" _
                        & Lf & "HTTP Response:  " & txt _
                        , Client.StatusCode
        End Select
        
        
        If txt = "null" Or txt = "[]" Then                      ' return blank collection
            Set IWebApiClient_SendRequest = New_c.Collection    '  if empty response
        
        ElseIf LenB(txt) <> 0 Then
            Set IWebApiClient_SendRequest = New_c.JSONDecodeToCollection(txt, parseErr)
            ErrorIf LenB(parseErr) <> 0, "Parse error:" & Lf & parseErr
        End If
        
    End With
ErrH: Blame "IWebApiClient_SendRequest"
End Function


Private Function GetRequestSignature(httpMethd As String _
                                   , contntMD5 As String _
                                   , reqstDate As Date _
                                   , reqstURL As String _
                                   ) As String
    Dim msgRepresentation$, secrtByts() As Byte, msgByts() As Byte, hmacByts() As Byte
    Const nl$ = vbLf
    On Error GoTo Cleanup
    
    msgRepresentation = UCase$(httpMethd) _
                 & nl & contntMD5 _
                 & nl & Format$(reqstDate, "mm/dd/yyyy hh:nn:ss") _
                 & nl & Client.UserName _
                 & nl & LCase$(reqstURL)
    
    With mCrypt
        msgByts = .VBStringToUTF8(msgRepresentation)
        secrtByts = .VBStringToUTF8(mPasswordHash)
        
        hmacByts = .HMAC_SHA1(msgByts, secrtByts, AsHexStr:=False)
        
        GetRequestSignature = mCrypt.Base64Enc(hmacByts)
    
    End With
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase secrtByts, msgByts, hmacByts
  LoadError "GetRequestSignature" ', "details of error"
End Function


Private Function GetContentMD5(reqstBody As String _
                             ) As String
    Dim md5Byts() As Byte, md5Str$
    On Error GoTo Cleanup
    
    If LenB(reqstBody) = 0 Then Exit Function
    
    md5Byts = mCrypt.VBStringToUTF8(reqstBody)
    md5Byts = mCrypt.MD5(md5Byts, AsHexStr:=False)
    md5Str = mCrypt.Base64Enc(md5Byts)
    
    GetContentMD5 = md5Str
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase md5Byts
  LoadError "GetContentMD5" ', "details of error"
End Function


Private Property Let IWebApiClient_Password(ByVal RHS As String)
    Dim byts() As Byte
    On Error GoTo Cleanup
    
    mPassword = RHS
    
    With mCrypt
        byts = .VBStringToUTF8(mPassword)
        byts = .SHA1(byts, AsHexStr:=False)
        
        mPasswordHash = mCrypt.Base64Enc(byts)
    
    End With
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase byts
  LoadError "IWebApiClient_Password" ', "details of error"
End Property






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

'Private Function Base64Encode(bytArray() As Byte) As String
'    Base64Encode = mCrypt.Base64Enc(bytArray)
'
'    If Right$(Base64Encode, 2) = vbCrLf Then _
'        Base64Encode = Left$(Base64Encode, Len(Base64Encode) - 2)
'
'End Function






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
    
    If mCrypt Is Nothing Then _
        Set mCrypt = New_c.Crypt
    
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
    
    Call Me.Cleanup
    
ErrH: Blame "Class_Terminate"
End Sub

Public Sub Cleanup()
    On Error GoTo ErrH
    
    Set mCrypt = Nothing
    Set mRequest = Nothing
    
ErrH: Blame "Cleanup"
End Sub






' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrRaise(errorMsg As String _
                   , Optional errorNumbr As TfiWebApiAuth1Errors = -1 _
                   , Optional errSource As String _
                   )
    Err.Raise errorNumbr, IIf(LenB(errSource) = 0, TypeName(Me), errSource), errorMsg
End Sub

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As TfiWebApiAuth1Errors = -1 _
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
                , Optional ByVal errorNumbr As TfiWebApiAuth1Errors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As TfiWebApiAuth1Errors = -1 _
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

