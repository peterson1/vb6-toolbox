VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "t_vbRichClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub Base64Enc_WontAppendVbCrLf()
    Dim s$
    With New_c.Crypt
        
        s = .Base64Enc("abc")
        
        Assert.That Right$(s, 2), Iz.Not.EqualTo(vbCrLf) _
              , "Last 2 chars shouldn't be linebreak."
        
'        If Right$(s, 2) = vbCrLf Then _
'            Debug.Print "String ends in vbCrLf" _
'                      , "Length: " & Len(s)
'
'        Debug.Print .Base64Dec(s) _
'                 , "Length: " & Len(.Base64Dec(s))
    End With
End Sub
