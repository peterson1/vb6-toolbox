VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWebApiClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum WebApiResponseCodes
    resp_OK = 200
    resp_Created = 201
    resp_BadRequest = 400
    resp_Unauthorized = 401
    resp_NotFound = 404
    resp_ServerError = 500
End Enum


Public UserName     As String
Public Password     As String

Public ContentType  As String
Public BaseURL      As String


Public Property Get ResponseText() As String
End Property





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Main Method
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Function SendRequest(httpMethd As String _
                          , url_withParams_withoutBase As String _
                          , Optional reqstBody As String _
                          , Optional raiseErrs As Boolean = True _
                          ) As cCollection
End Function




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructor
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Sub Init(wbApiBaseUrl As String _
              , Optional userNme As String _
              , Optional passwrd As String _
              , Optional contentTyp$ = "application/json; charset=utf-8" _
              )
End Sub




' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Status Accessors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Public Property Get StatusCode() As WebApiResponseCodes
End Property

Public Property Get StatusText() As String
End Property

