Attribute VB_Name = "MMain"
#If TWINBASIC Then
Module Main
#End If

'Split a string into a new array.
'Returns the number of elements in the array.
'If lMaxItems is specified, then the returned asRetItems() array
'will have at maximum lMaxItems, the last one holding the remaining
'chunk that wasn't splitted because the lMaxItems limit was reached.
Public Function SplitString(ByRef asRetItems() As String, _
  ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As Long

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim lRetCount   As Long
  
  On Error Resume Next
  Erase asRetItems
  On Error GoTo SplitString_Err
  
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        lRetCount = lRetCount + 1&
        ReDim Preserve asRetItems(1& To lRetCount)
        asRetItems(lRetCount) = Left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If lRetCount = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    lRetCount = lRetCount + 1&
    ReDim Preserve asRetItems(1& To lRetCount)
    asRetItems(lRetCount) = sToSplit
  End If
  SplitString = lRetCount
SplitString_Err:
End Function

'Function needed as vb invokes default member on VarType(A), where A is a CRow/CList
Public Function GetVarType(ByRef pvValue As Variant) As VbVarType
  Dim iType As VbVarType
  On Error Resume Next
  Err.Clear
  If IsObject(pvValue) Then
    iType = vbObject
  Else
    iType = VarType(pvValue)
  End If
  GetVarType = iType
End Function

#If (MSACCESS = 0) And (TWINBASIC = 0) Then
Public Sub Main()
  RunAllTests
End Sub

Public Function Nz(ByRef pvValue As Variant, ByVal pvDefault As Variant) As Variant
  If Not IsObject(pvValue) Then
    If Not IsNull(pvValue) Then
      Nz = pvValue
    Else
      Nz = pvDefault
    End If
  Else
    If pvValue Is Nothing Then
      If Not IsObject(pvDefault) Then
        Nz = pvDefault
      Else
        Set Nz = pvDefault
      End If
    Else
      Set Nz = pvValue
    End If
  End If
End Function

Public Function Replace(ByVal sIn As String, ByVal sFind As _
    String, ByVal sReplace As String, Optional nStart As _
     Long = 1, Optional nCount As Long = -1, _
     Optional Compare As VbCompareMethod = vbBinaryCompare) As _
     String

  Dim nC As Long, nPos As Long
  Dim nFindLen As Long, nReplaceLen As Long

  nFindLen = Len(sFind)
  nReplaceLen = Len(sReplace)
  
  If (sFind <> "") And (sFind <> sReplace) Then
    nPos = InStr(nStart, sIn, sFind, Compare)
    Do While nPos
        nC = nC + 1
        sIn = Left(sIn, nPos - 1) & sReplace & _
         Mid(sIn, nPos + nFindLen)
        If nCount <> -1 And nC >= nCount Then Exit Do
        nPos = InStr(nPos + nReplaceLen, sIn, sFind, _
          Compare)
    Loop
  End If

  Replace = sIn
End Function

Public Function Split(ByVal sIn As String, _
  Optional Delim As String = " ", _
  Optional Limit As Long = -1, _
  Optional Compare As VbCompareMethod = vbBinaryCompare) _
  As Variant

  Dim nC As Long, nPos As Long, nDelimLen As Long
  Dim sOut() As String
  
  If Delim <> "" Then
    nDelimLen = Len(Delim)
    nPos = InStr(1, sIn, Delim, Compare)
    Do While nPos
      nC = nC + 1
      ReDim Preserve sOut(1 To nC)
      sOut(nC) = Left(sIn, nPos - 1)
      sIn = Mid(sIn, nPos + nDelimLen)
      If Limit <> -1 And nC >= Limit Then Exit Do
      nPos = InStr(1, sIn, Delim, Compare)
    Loop
  End If

  ReDim Preserve sOut(1 To nC)
  sOut(nC) = sIn

  Split = sOut
End Function

#End If

#If TWINBASIC Then
End Module
#End If
