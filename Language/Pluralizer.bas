Attribute VB_Name = "Plu"
Option Explicit

Private mSingular$

'  usage:
'
'    Plu.ral("apple")               <-- returns "apples"
'    Plu.ral("glass", 8)            <-- returns "8 glasses"
'    Plu.ral("file", 0) & " found"  <-- returns "No files found"
'
Public Property Get ral(singulrNoun As String _
                      , Optional quantty As Long = -1 _
                      , Optional use_No_ifZero As Boolean = True _
                      ) As String
    
    Dim plurlForm$: plurlForm = PluralForm(singulrNoun)
    
    Select Case quantty
        
        Case -1     ' do NOT include quantity
            ral = plurlForm
            
        
        Case 0      ' append plural form to "No" or "0"
            ral = IIf(use_No_ifZero, "No ", "0 ") & plurlForm
        
        
        Case 1      ' use singular form
            ral = "1 " & singulrNoun
        
        
        Case Else   ' append plural form to formatted number
            ral = Format$(quantty, "#,# ") & plurlForm
        
    End Select
End Property



Public Property Get PluralForm(singulrNoun As String) As String
    
    Select Case LCase$(singulrNoun)
        
        '  except proper nouns
        Case "nationwide", "nestle", "getz"
            PluralForm = singulrNoun
        
        '  except other nouns
        Case "all"
            PluralForm = singulrNoun
        
        Case Else
            PluralForm = EnglishPlural(singulrNoun)
            
    End Select
End Property


Private Function EnglishPlural(singulrNoun As String) As String
    mSingular = singulrNoun
    
    If Ends("s") Then
        EnglishPlural = ChangeEndTo("s", "es")
    
    ElseIf Ends("y") Then
        EnglishPlural = ChangeEndTo("i", "es")
        
    ElseIf Ends("fe") Then
        EnglishPlural = ChangeEndTo("ve", "s")
        
    Else
        EnglishPlural = ChangeEndTo("", "s")
    End If
     
End Function


Private Function Ends(lastChars$) As Boolean
    Ends = UCase$(Right$(mSingular, Len(lastChars))) = UCase$(lastChars)
End Function

Private Function ChangeEndTo(replacemnt As String _
                           , suffx As String _
                           ) As String
    
    ChangeEndTo = Left$(mSingular, Len(mSingular) _
                                 - Len(replacemnt)) _
                & IIf(IsUppercase(Right$(mSingular, 1)) _
                    , UCase$(replacemnt & suffx) _
                           , replacemnt & suffx)
End Function

Private Function IsUppercase(charactr$) As Boolean
    Dim a%: a = Asc(charactr)
              ' Asc(A)=65    Asc(Z)=90
    IsUppercase = (a > 64) And (a < 91)
    
End Function

























