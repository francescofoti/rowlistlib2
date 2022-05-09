Attribute VB_Name = "MRowList"
#If TWINBASIC Then
Module MRowList
#End If

#If MSACCESS Then
Option Compare Database
#End If

Option Explicit

'Returns an array with 4 entries which are the lower and upper bounds of
'the first two dimensions of the array stored in pvVar.
'The number of dimensions is returned in piRetDims.
Public Sub GetVarArrayBounds(ByRef pvVar As Variant, ByRef piRetDims As Integer, ByRef pavRetBounds As Variant)
  Dim lLB1    As Long
  Dim lUB1    As Long
  Dim lLB2    As Long
  Dim lUB2    As Long
  
  On Error Resume Next
  piRetDims = 0
  lLB1 = LBound(pvVar)
  If Err.Number = 0 Then
    piRetDims = 1
    lUB1 = UBound(pvVar)
    lLB2 = LBound(pvVar, 2)
    If Err.Number = 0& Then
      piRetDims = 2
      lUB2 = UBound(pvVar, 2)
    End If
  End If
  pavRetBounds = Array(lLB1, lUB1, lLB2, lUB2)
End Sub

' Build a row from a serie of values; all columns will be named "#"+<column number>.
' Build a row from a serie of values; all columns will be named "#"+<column number>.
Public Function MakeRow(ParamArray pavValues() As Variant) As CRow
  Dim lLowerBound   As Long
  Dim lUpperBound   As Long
  Dim i             As Long
  Dim iCol          As Long
  Dim rowValues     As CRow
  
  'get the number of elements in pavValues
  On Error Resume Next
  lLowerBound = LBound(pavValues())
  If Err.Number = 0& Then 'If we could get the lBound then we have at least one element
    lUpperBound = UBound(pavValues())
    Set rowValues = New CRow
    iCol = 1&
    For i = lLowerBound To lUpperBound
      rowValues.AddCol "#" & iCol, pavValues(i), 0&, 0&
      iCol = iCol + 1&
    Next i
    Set MakeRow = rowValues
  End If
End Function
  
'
' Utilities for lists
'

'Creates a list with two columns, "ParamName" and "Value" and
'adds the pairs of elements in pavPairs() in each respective
'column, expecting the first element of the pair to be the
'"ParamName" column value, and the second one being the "Value" column value.
'If the number of parameters is not even, the function does not
'create a list and returns nothing.
Public Function MakeParamList(ParamArray pavPairs() As Variant) As CList
  Dim lLowerBound   As Long
  Dim lElemCount    As Long
  Dim lPairsCount   As Long
  Dim lPairIndex    As Long
  Dim lstRetPairs   As CList
  
  'get the number of elements in pavPairs
  On Error Resume Next
  lLowerBound = LBound(pavPairs())
  If Err.Number = 0& Then 'If we could get the lBound then we have at least one element
    lElemCount = UBound(pavPairs()) - lLowerBound + 1& 'Just in case 0 is not the lower bound
    If (lElemCount Mod 2&) = 0& Then 'We must have pairs (name, followed by value)
      Set lstRetPairs = New CList
      lstRetPairs.ArrayDefine Array("ParamName", "Value"), Array(vbString, vbVariant)
      lPairsCount = lElemCount \ 2&
      For lPairIndex = 1& To lPairsCount
        lstRetPairs.AddValues pavPairs((lLowerBound + (lPairIndex - 1&)) * 2&), _
                              pavPairs((lLowerBound + (lPairIndex - 1&)) * 2& + 1&)
      Next lPairIndex
      Set MakeParamList = lstRetPairs
    Else
      Set MakeParamList = Nothing
    End If
  End If
End Function

Public Function MakeParamRow(ParamArray pavValues() As Variant) As CRow
  Dim lLowerBound   As Long
  Dim lUpperBound   As Long
  Dim i             As Long
  Dim iCol          As Long
  Dim rowValues     As CRow
  
  'get the number of elements in pavValues
  On Error Resume Next
  lLowerBound = LBound(pavValues())
  If Err.Number = 0& Then 'If we could get the lBound then we have at least one element
    lUpperBound = UBound(pavValues())
    Set rowValues = New CRow
    iCol = 1&
    For i = lLowerBound To lUpperBound
      rowValues.AddCol "#" & iCol, pavValues(i), 0&, 0&
      iCol = iCol + 1&
    Next i
    Set MakeParamRow = rowValues
  End If
End Function

'Split a string and insert elements in a list.
'Returns a new list object that must be freed setting it to nothing.
Public Function SplitToList(ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As CList

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim oRetList    As CList
  
  On Error GoTo SplitToList_Err
  
  Set oRetList = New CList
  oRetList.Define "Item", vbString, 0&, 0&
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        oRetList.AddValues Left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If oRetList.Count = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    oRetList.AddValues sToSplit
  End If
  Set SplitToList = oRetList
  Exit Function

SplitToList_Err:
  Set oRetList = Nothing
End Function

Public Function JoinList(ByRef lstToJoin As CList, Optional ByVal sColSep As String = ",", Optional ByVal sRowSep As String = vbCrLf, Optional ByVal psColFilter As String = "") As String
  Dim iCol        As Long
  Dim iRow        As Long
  Dim sRet        As String
  Dim asColName() As String
  Dim iColCt      As Long
  
  On Error Resume Next
  If Len(psColFilter) Then
    iColCt = SplitString(asColName(), psColFilter, ";")
  End If
  For iRow = 1& To lstToJoin.Count
    If iRow > 1& Then sRet = sRet & sRowSep
    If iColCt = 0& Then
      For iCol = 1& To lstToJoin.ColCount
        If iCol > 1& Then sRet = sRet & sColSep
        sRet = sRet & lstToJoin(iCol, iRow) & ""
      Next iCol
    Else
      For iCol = 1& To iColCt
        If iCol > 1& Then sRet = sRet & sColSep
        sRet = sRet & lstToJoin(asColName(iCol), iRow) & ""
      Next iCol
    End If
  Next iRow
  JoinList = sRet
End Function
  
Public Function StrBlock(ByVal sText As String, ByVal sPadChar As String, ByVal iMaxLen As Integer) As String
  Dim iLen      As Integer
  
  iLen = Len(sText)
  If iLen < iMaxLen Then
    StrBlock = sText & String$(iMaxLen - iLen, sPadChar)
  Else
    If iMaxLen > 6 Then
      StrBlock = Left$(sText, iMaxLen - 3) & "..."
    Else
      StrBlock = Left$(sText, iMaxLen)
    End If
  End If
End Function

'Handle object and null values to go from variant to string.
'object --> "#ref"
'null --> "#null"
'empty --> "#empty"
Public Function VariantAsString(ByRef pvValue As Variant, Optional ByVal pfHexNumbers As Boolean = False) As String
  Dim vType         As VbVarType
  
  If IsObject(pvValue) Then
    If Not pvValue Is Nothing Then
      VariantAsString = "#ref_" & TypeName(pvValue) & "_" & ObjPtr(pvValue)
    Else
      VariantAsString = "#nothing"
    End If
  Else
    vType = VarType(pvValue)
    If vType = vbNull Then
      VariantAsString = "#null"
    ElseIf vType = vbEmpty Then
      VariantAsString = "#empty"
    Else
      Select Case vType
      Case vbInteger, vbLong, vbByte, vbDouble, vbCurrency, vbDecimal
        If pfHexNumbers Then
          VariantAsString = "$" & LCase$(Hex$(pvValue))
        Else
          VariantAsString = pvValue & ""
        End If
      Case Else
        VariantAsString = pvValue & ""
      End Select
    End If
  End If
End Function

Private Sub Indent(ByVal piLevel As Integer)
  If piLevel > 0 Then
    Output Space$(piLevel * 2)
  End If
End Sub

Private Sub ListDumpHeader( _
  ByRef poList As CList, _
  ByRef paiColWidth() As Integer, _
  Optional ByVal piLevel As Integer = 0, _
  Optional ByVal psTitle As String = "", _
  Optional ByVal psColWidths As String = "", _
  Optional ByVal pfShowColTitles As Boolean = True)
  Dim lCount        As Long
  Dim i             As Integer
  Dim k             As Integer
  Dim iLen          As Integer
  
  Dim iColWidthCt       As Integer
  Dim asColWidthSpec()  As String
  Dim iCol              As Integer
  Dim sColName          As String
  Dim sWidth            As String
  Dim iColon            As Integer

  lCount = poList.ColCount
  ReDim paiColWidth(1 To lCount) As Integer
  Indent piLevel
  If Len(psTitle) Then
    OutputLn String$(Len(psTitle), "-") & "+"
    Indent piLevel
    OutputLn psTitle & "|"
  End If

  If Len(psColWidths) Then
    iColWidthCt = SplitString(asColWidthSpec(), psColWidths, ";")
    For i = 1 To iColWidthCt
      iColon = InStr(1, asColWidthSpec(i), ":")
      If iColon Then
        sColName = Left$(asColWidthSpec(i), iColon - 1)
        sWidth = Right$(asColWidthSpec(i), Len(asColWidthSpec(i)) - iColon)
        If Len(sWidth) > 0 Then
          If Val(sWidth) > 0 Then
            If InStr(1, sColName, "*") = 0 Then
              iCol = poList.ColPos(sColName)
              If iCol Then
                paiColWidth(iCol) = Val(sWidth)
              End If
            Else
              For k = 1 To lCount
                If poList.ColName(k) Like sColName Then
                  paiColWidth(k) = Val(sWidth)
                End If
              Next k
            End If
          End If
        End If
      End If
    Next i
  End If
  
  If pfShowColTitles Then
    Indent piLevel

    'col titles row sep
    For i = 1 To lCount
      iLen = IIf(paiColWidth(i) = 0, Len(poList.ColName(i)), paiColWidth(i))
      Output String$(iLen, "-") & "+"
    Next i
    OutputLn ""

    Indent piLevel
    For i = 1 To lCount
      iLen = IIf(paiColWidth(i) = 0, Len(poList.ColName(i)), paiColWidth(i))
      Output StrBlock(poList.ColName(i), " ", iLen) & "|"
    Next i
    OutputLn ""

    'col titles row sep
    Indent piLevel
    For i = 1 To lCount
      iLen = IIf(paiColWidth(i) = 0, Len(poList.ColName(i)), paiColWidth(i))
      Output String$(iLen, "-") & "+"
    Next i
    OutputLn ""
  End If
  
End Sub

Private Sub ListDumpHeaderHTML( _
  poList As CList, _
  Optional ByVal piLevel As Integer = 0, _
  Optional ByVal psTitle As String = "", _
  Optional ByVal pfShowColTitles As Boolean = True)
  
  Dim lCount        As Long
  Dim i             As Integer
  Dim k             As Integer
  Dim iLen          As Integer
  
  Dim iColWidthCt       As Integer
  Dim asColWidthSpec()  As String
  Dim iCol              As Integer
  Dim sColName          As String
  Dim sWidth            As String
  Dim iColon            As Integer

  Output "<thead>"
  
  lCount = poList.ColCount
  If Len(psTitle) Then
    Output "<caption>" & psTitle & "</caption>"
  End If

  If pfShowColTitles Then
    Output "<tr>"
    For i = 1 To lCount
      Output "<th>" & poList.ColName(i) & "</th>"
    Next i
    Output "</tr>"
  End If

  Output "</thead>"
End Sub
  
Public Sub ListDump( _
    poRowList As Object, _
    Optional ByVal sTitle As String = "", _
    Optional ByVal psColWidths As String = "", _
    Optional ByVal plStartRow As Long = 0&, _
    Optional ByVal plEndRow As Long = 0&, _
    Optional ByVal pfShowColTitles As Boolean = True, _
    Optional ByVal pfHexIntLongs As Boolean = False, _
    Optional ByVal piLevel As Integer = 0, _
    Optional ByVal pfDeepDump As Boolean = False, _
    Optional ByVal pfHTML As Boolean = False, _
    Optional ByVal pfIgnoreUnattended As Boolean = False)
  Dim iRow      As Long
  Dim i         As Long
  Dim k         As Long
  Dim lCount    As Long
  Dim asColName()  As String
  Dim iLen      As Integer
  Dim iStart    As Long
  Dim iEnd      As Long
  Dim oList     As CList
  Dim oRow      As CRow
  Dim aiColWidth()      As Integer
  
  Dim oSubObject        As Object
  Dim vValue            As Variant
  
  On Error GoTo ListDump_Err
  
  #If TESTING Then
    If Nz(Test_GetParam(SUITEPARAM_LISTDUMPOFF), False) Then
      If Not pfIgnoreUnattended Then
        Exit Sub
      End If
    End If
  #End If
  
  'Handle row or list parameter
  If poRowList Is Nothing Then Exit Sub
  If TypeOf poRowList Is CList Then
    Set oList = poRowList
  ElseIf TypeOf poRowList Is CRow Then
    'We create a list from the row
    Set oList = New CList
    Set oRow = poRowList
    oRow.DefineList oList
    oList.AddRow oRow
  Else
    Debug.Print "ListDump() Invalid Parameter #1 (Class: '" & TypeName(poRowList) & "'): Class Must be CRow or CList."
    Exit Sub
  End If
  
  lCount = oList.ColCount
  If lCount = 0& Then Exit Sub
  
  ' HEADER
  '--------
  If Not pfHTML Then
    ListDumpHeader oList, aiColWidth, piLevel, sTitle, psColWidths, pfShowColTitles
  Else
    Output "<table border=""1"">"
    If (TypeOf poRowList Is CList) And (oList.ColCount = 1) Then
      Debug.Print "ColCount=" & oList.ColCount & ", colname1=" & oList.ColName(1)
    Else
      ListDumpHeaderHTML oList, piLevel, sTitle, pfShowColTitles
    End If
  End If

  'dump values
  '---------------------------------------------
  iStart = 1&
  iEnd = oList.Count
  If plStartRow > 0& Then
    If plStartRow <= oList.Count Then
      iStart = plStartRow
    End If
  End If
  If plEndRow > 0& Then
    If plEndRow <= oList.Count Then
      iEnd = plEndRow
    End If
  End If
  If iStart > iEnd Then
    'Swap
    Dim iTemp As Long
    iTemp = iEnd
    iEnd = iStart
    iStart = iTemp
  End If
  
  ReDim afIsColValueObject(1 To lCount) As Boolean
  
  If pfHTML Then
    Output "<tbody>"
  End If
  
  For iRow = iStart To iEnd
    If pfHTML Then
      Output "<tr>"
    Else
      Indent piLevel
    End If
    
    For i = 1 To lCount
      afIsColValueObject(i) = False
      If oList.IsItemObject(i, iRow) Then
        afIsColValueObject(i) = True
        Set vValue = oList(i, iRow)
      Else
        vValue = oList(i, iRow)
      End If
      
      If Not pfHTML Then
        iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
        Output StrBlock(VariantAsString(vValue, pfHexIntLongs) & "", " ", iLen) & "|"
      Else
        Output "<td>" & ToHTML(VariantAsString(vValue, pfHexIntLongs))
        If afIsColValueObject(i) Then
          If pfDeepDump Then
            Set oSubObject = vValue
            If Not oSubObject Is Nothing Then
              If (TypeOf oSubObject Is CRow) Or (TypeOf oSubObject Is CList) Then
                sTitle = TypeName(oSubObject) & " @col,row[" & i & ", " & iRow & "]"
                ListDump oSubObject, sTitle, _
                        psColWidths, _
                        pfShowColTitles:=pfShowColTitles, _
                        pfHexIntLongs:=pfHexIntLongs, _
                        piLevel:=piLevel + 1, _
                        pfDeepDump:=pfDeepDump, _
                        pfHTML:=pfHTML, pfIgnoreUnattended:=pfIgnoreUnattended
              End If
            End If
            Set oSubObject = Nothing
          End If
        End If
      End If
      If afIsColValueObject(i) Then
        Set vValue = Nothing
      End If
      If pfHTML Then
        Output "</td>"
      End If
    Next i

    If Not pfHTML Then
      OutputLn ""
    Else
      Output "</tr>"
    End If

    'Output sub objects
    If pfDeepDump And (Not pfHTML) Then
      For i = 1 To lCount
        If afIsColValueObject(i) Then
          Set oSubObject = oList(i, iRow)
          If Not oSubObject Is Nothing Then
            If (TypeOf oSubObject Is CRow) Or (TypeOf oSubObject Is CList) Then
              sTitle = TypeName(oSubObject) & " @col,row[" & i & ", " & iRow & "]"
              ListDump oSubObject, sTitle, _
                      psColWidths, _
                      pfShowColTitles:=pfShowColTitles, _
                      pfHexIntLongs:=pfHexIntLongs, _
                      piLevel:=piLevel + 1, _
                      pfDeepDump:=pfDeepDump, _
                      pfHTML:=pfHTML, pfIgnoreUnattended:=pfIgnoreUnattended
            End If
          End If
        End If
        Set oSubObject = Nothing
      Next i
    End If
    
  Next iRow
  If pfHTML Then
    OutputLn "</tbody></table>"
  End If

ListDump_Exit:
  On Error Resume Next
  If IsObject(vValue) Then
    Set vValue = Nothing
  End If
  Set oSubObject = Nothing
  Set oRow = Nothing
  Set oList = Nothing
  Exit Sub
  
ListDump_Err:
  Resume ListDump_Exit
  Resume
End Sub

#If TWINBASIC Then
End Module
#End If

