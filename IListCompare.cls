VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If TWINBASIC Then
Class IListCompare
#End If

'Read-only parameters that can (or not) be object references, although are references.
'Return same values as StrComp().
Public Function CompareValues(ByVal plColIndex As Long, ByVal pvValue1 As Variant, ByVal pvvalue2 As Variant) As Integer

End Function

'second methods, so we can take into account woldcard/joker search as CList does
Public Function CompareStrings(ByVal plColIndex As Long, ByVal pvString1 As Variant, ByVal pvString2 As Variant, ByVal peCompMethod As VbCompareMethod) As Integer

End Function
  
#If TWINBASIC Then
End Class
#End If
