Attribute VB_Name = "MJsonRowListHelper"
#If TWINBASIC Then
Module MJsonRowListHelper
#End If
  
'Try to follow a path in a json object parsed in CRow/CList objects.
'pavPath is an array in a variant, where each element can be
' - A string: json key to follow (name of the CRow/CList cell where
'   we find either the next object reference (CRow/CList), or
' - An array in a Variant (array of two elements, with LBound=0).
'   Meaning we should be on a CList object on our path.
'   The array contains the coordinates where to find the
'   value or next object ref we're looking for:
'   - 1. The column name (string) or index (long)
'   - 2. The line number (long)
Public Function FollowPath(ByRef poJson As Object, ByRef pavPath As Variant) As String
  Const LOCAL_ERR_CTX As String = "FollowPath"
  
  Dim oObject   As Object
  Dim avPath    As Variant
  Dim iPathCt   As Integer
  Dim i         As Integer
  Dim counter   As Integer
  Dim sPathTrail As String
  
  On Error GoTo FollowPath_Err
  
  Set oObject = poJson
  iPathCt = UBound(pavPath) - LBound(pavPath) + 1
  For i = LBound(pavPath) To UBound(pavPath)
    avPath = pavPath(i)
    counter = counter + 1
    If counter < iPathCt Then
      If Not IsArray(avPath) Then
        sPathTrail = CombinePath(sPathTrail, avPath)
        Set oObject = oObject.ColValue(CStr(avPath))
      Else
        sPathTrail = CombinePath(sPathTrail, avPath(0) & "," & avPath(1))
        Set oObject = oObject.Item(avPath(0), CLng(avPath(1)))
      End If
    Else
      If Not IsArray(avPath) Then
        FollowPath = oObject.ColValue(CStr(avPath))
      Else
        FollowPath = oObject.Item(CStr(avPath(0)), CLng(avPath(1)))
      End If
    End If
  Next

FollowPath_Exit:
  Set oObject = Nothing
  Exit Function
  
FollowPath_Err:
  Err.Raise LOCAL_ERR_CTX, Err.Number, "FollowPath Failed, trail=[" & sPathTrail & "]: " & Err.Description
  Resume FollowPath_Exit
End Function

#If TWINBASIC Then
End Module
#End If
