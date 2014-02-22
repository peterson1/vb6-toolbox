Attribute VB_Name = "Tag_"
Option Explicit

Private mXML As New XmlConstraints


Public Function NameIs(nameOfTag As String _
                     ) As IConstraint
    Call mXML.Init(fn_NameIs, nameOfTag)
    Set NameIs = mXML
End Function


Public Function HasTag(nameOfTag As String _
                     ) As IConstraint
    Call mXML.Init(fn_HasTag, nameOfTag)
    Set HasTag = mXML
End Function


Public Function ValueIs(attrbuteValue As Variant _
                      ) As IConstraint
    Call mXML.Init(fn_ValueIs, attrbuteValue)
    Set ValueIs = mXML
End Function


Public Function TextIs(textOfTag As String _
                     ) As IConstraint
    Call mXML.Init(fn_TextIs, textOfTag)
    Set TextIs = mXML
End Function


Public Function Tag(nameOfTag As String _
                  ) As XmlConstraints
    mXML.ParentTag = nameOfTag
    Set Tag = mXML
End Function

Public Function Find(searchFiltr As String _
                   ) As XmlConstraints
    mXML.SearchFilter = searchFiltr
    Set Find = mXML
End Function
