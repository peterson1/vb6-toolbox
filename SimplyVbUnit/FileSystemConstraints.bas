Attribute VB_Name = "Path"
Option Explicit

Private mFS As New FileSystemConstraints


Public Function DoesNotExist() As IConstraint

    Call mFS.Init(fn_DoesNotExist)
    
    Set DoesNotExist = mFS
    
End Function


Public Function Exists() As IConstraint
    
    Call mFS.Init(fn_Exists)
    
    Set Exists = mFS
    
End Function
