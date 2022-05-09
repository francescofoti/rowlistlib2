Attribute VB_Name = "MRowTests"
#If TWINBASIC Then
Module MRowTests
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Public Sub Test_RunAllTests()
  On Error GoTo Test_RunAllTests_Err
  
  Test_BeginSuite "MRowTests"
  
  TestCollectionVSMapStringToLong
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestCollectionKeysCaseInsensitive
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestMapDuplicateKeys
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestPopulateWithAddCol
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestRowArrayDefine
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestRowDefine
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestArrayAssign
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestCopyAndCloningRows
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestRowErrors
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestMergeRows
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit
  
  TestSharpNotation
  If Not Test_LastSuiteSuccess() Then GoTo Test_RunAllTests_Exit

Test_RunAllTests_Exit:
  Test_EndSuite
  Exit Sub

Test_RunAllTests_Err:
  Debug.Print "Test_RunAllTests() failed: " & Err.Description
  Resume Test_RunAllTests_Exit
  Resume
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestCollectionVSMapStringToLong()
  Const LOCAL_ERR_CTX As String = "TestCollectionVSMapStringToLong"
  Call OpenTraceOutputFile
  
  On Error GoTo T1_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Compare performance of a collection against the performance " & _
               "of the CMapStringToLong class"

  Dim cCollection   As Collection
  Dim mslMap        As CMapStringToLong
  Dim i             As Long
  Dim j             As Long
  Dim k             As Long
  Dim lRandom       As Long
  Dim lValue        As Long
  
  Set cCollection = New Collection
  Set mslMap = New CMapStringToLong
  
  '
  ' Adding a huge number of elements
  '
  'Const MAX_ELEMENTS As Long = 100000
  Const MAX_ELEMENTS As Long = 50000
  
  'Collection
  Dim iPercent As Integer
  Dim iLastPercent As Integer
  
  Output "Collection: Adding " & MAX_ELEMENTS & "... "
  ChronoStart
  For i = 1 To MAX_ELEMENTS
    cCollection.Add i, CStr(i)
'    iPercent = CInt((i * 100) / MAX_ELEMENTS)
'    If iPercent > 0 Then
'      If iPercent <> iLastPercent Then
'        Debug.Print "Added #" & i & " elements ..." & iPercent & "%"
'        iLastPercent = iPercent
'      End If
'    End If
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Adding " & MAX_ELEMENTS & "... "
  ChronoStart
  For i = 1 To MAX_ELEMENTS
    mslMap.Add CStr(i), i
  Next i
  mslMap.Sorted = True
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Retrieval #1: by numerical index
  '
  Const MAX_TEST_ELEMENTS As Long = 1000&
  Const MAX_SHUFFLES      As Long = 500&
  
  Dim e1 As Long
  Dim e2 As Long
  Dim xx As Integer
  Dim alIndex(1 To MAX_TEST_ELEMENTS) As Long
  'Get random elements
  OutputLn "Shuffling " & MAX_TEST_ELEMENTS & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS
    alIndex(i) = i
  Next i
  xx = 0
  For i = 1 To MAX_SHUFFLES * 4
    Do
      e1 = GetRandom(1&, MAX_TEST_ELEMENTS)
      e2 = GetRandom(1&, MAX_TEST_ELEMENTS)
    Loop Until e1 <> e2
    k = alIndex(e1)
    alIndex(e1) = alIndex(e2)
    alIndex(e2) = k
    
    If (i Mod 50) = 0 Then
      Debug.Print ".";
      xx = xx + 1
    End If
    If xx = 50 Then
      Debug.Print " " & i & "/" & MAX_SHUFFLES & " shuffles"
      xx = 0
    End If
  Next i
  Debug.Print
  OutputLn ChronoTime() & " seconds."
  
  'Collection
  Const MAX_ERRORS As Long = 10&
  Dim lErrCt As Long
  
  Output "Collection: Retrieving (by numerical index) " & MAX_TEST_ELEMENTS & " elements... "
  ChronoStart
  On Error Resume Next
  For i = 1 To MAX_TEST_ELEMENTS
    lValue = cCollection(alIndex(i))
    If Err.Number <> 0 Then
      OutputLn "Error at loop #" & i & " accessing collection element #" & alIndex(i)
      lErrCt = lErrCt + 1
      If lErrCt > MAX_ERRORS Then
        Exit For
      End If
    End If
    DoEvents
  Next i
  On Error GoTo T1_Err
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Retrieving (by numerical index) " & MAX_TEST_ELEMENTS & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS
    lValue = mslMap.Item(alIndex(i))
  Next i
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Retrieval #2: by alphabetical index
  '
  'Collection
  Output "Collection: Retrieving (by key) " & MAX_TEST_ELEMENTS & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS
    lValue = cCollection(CStr(alIndex(i)))
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Retrieving (by key) " & MAX_TEST_ELEMENTS & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS
    lValue = mslMap.Item(mslMap.Find(CStr(alIndex(i))))
  Next i
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Removing elements
  '
  'Get random elements again, but take them at the beginning of the
  'full range, otherwise, there is the risk that the element index
  'that will be removed is out of the valid bounds.
  Output "Sorting " & MAX_TEST_ELEMENTS & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS \ 100
    lRandom = GetRandom(1&, MAX_ELEMENTS \ 100)
    alIndex(i) = lRandom
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'Collection
  Output "Collection: Removing " & (MAX_TEST_ELEMENTS \ 100) & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS \ 100
    cCollection.Remove alIndex(i)
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Removing " & (MAX_TEST_ELEMENTS \ 100) & " elements... "
  ChronoStart
  For i = 1 To MAX_TEST_ELEMENTS \ 100
    mslMap.Remove alIndex(i)
  Next i
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Destruction
  '
  Output "Collection: destroy... "
  ChronoStart
  Set cCollection = Nothing
  OutputLn ChronoTime() & " seconds."
  
  Output "CMapStringToLong: destroy... "
  ChronoStart
  Set mslMap = Nothing
  OutputLn ChronoTime() & " seconds."
  
  OutputLn "SUCCESS"
  Test_SetSuccess LOCAL_ERR_CTX, True, LOCAL_ERR_CTX & " success."
  
T1_Exit:
  On Error Resume Next
  CloseTraceOutputFile
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

T1_Err:
  OutputLn "FATAL Error in test method: " & Err.Description
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume T1_Exit
  Resume
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestCollectionKeysCaseInsensitive()
  Const LOCAL_ERR_CTX As String = "TestCollectionKeysCaseInsensitive"
  On Error GoTo TestCollectionKeysCaseInsensitive_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Prove that collection keys are case insensitive"
  
  'Prove that collection keys are case insensitive
  Dim cCollection As New Collection
  cCollection.Add Item:="Item1", Key:="iTeM1"
  On Error Resume Next
  cCollection.Add Item:="Item1", Key:="item1" 'this generates a runtime error
  If Err.Number Then
    OutputLn "Error #" & Err.Number & " occured while trying to add a duplicate key (but with different letter case) in a collection."
    Test_SetSuccess LOCAL_ERR_CTX, True, LOCAL_ERR_CTX & " success."
  Else
    Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: Collection keys ARE case sensitive"
  End If
  
TestCollectionKeysCaseInsensitive_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestCollectionKeysCaseInsensitive_Err:
  OutputLn LOCAL_ERR_CTX & " error: " & Err.Description
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestCollectionKeysCaseInsensitive_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestMapDuplicateKeys()
  Const LOCAL_ERR_CTX As String = "TestMapDuplicateKeys"
  
  On Error GoTo TestMapDuplicateKeys_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Add some empty and duplicate keys to a CMapStringToInteger object, " & _
               "then remove duplicates."
  
  Dim mslMap    As New CMapStringToLong
  Dim i         As Long
  Dim k         As Long
  Dim iCount    As Integer
  
  Const ABC_OCCUR_COUNT = 5 'number of 'abc' items will add
  Const UNIQUE_ITEMS_COUNT = 3
  'Add to the map, take care to stay in sync with previous consts
  mslMap.Add "", 1&
  mslMap.Add "", 2&
  mslMap.Add "", 3&
  mslMap.Add "abc", 10&
  mslMap.Add "abc", 100&
  mslMap.Add "def", 20&
  mslMap.Add "def", 200&
  mslMap.Add "abc", 101&
  mslMap.Add "abc", 102&
  mslMap.Add "abc", 103&
  mslMap.Sorted = True
  
  OutputLn mslMap.Count & " items in set."
  i = mslMap.Find("abc")
  OutputLn "one of the 'abc' found at position: " & i
  Test_ValueNotEqual 0, i, "'abc' item not found"
  
  'print all items which key is "abc"
  OutputLn "All items which key is 'abc': "
  k = mslMap.FindFirst("abc")
  If k Then
    For i = k To mslMap.Count
      If mslMap.Key(i) = "abc" Then
        OutputLn i & ": " & mslMap.Key(i) & " --> " & mslMap.Item(i)
        iCount = iCount + 1
      Else
        Exit For
      End If
    Next i
  Else
    OutputLn "key 'abc' was not found by the FindFirst Function() ???"
  End If
  Test_Value ABC_OCCUR_COUNT, iCount, "exactly " & ABC_OCCUR_COUNT & "'abc' item(s) should have been found"

  OutputLn "Removing duplicates..."
  mslMap.RemoveDuplicates
  OutputLn mslMap.Count & " items remaining in set: "
  For i = 1 To mslMap.Count
    OutputLn i & ": " & mslMap.Key(i) & " --> " & mslMap.Item(i)
  Next i
  Test_Value UNIQUE_ITEMS_COUNT, mslMap.Count, "exactly " & UNIQUE_ITEMS_COUNT & " unique item(s) should have been counted"

TestMapDuplicateKeys_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestMapDuplicateKeys_Err:
  OutputLn LOCAL_ERR_CTX & " error: " & Err.Description
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestMapDuplicateKeys_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestPopulateWithAddCol()
  Const LOCAL_ERR_CTX As String = "TestPopulateWithAddCol"
  On Error GoTo TestPopulateWithAddCol_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Populate column set of a CRow object using the AddCol method"

  Dim oRow          As New CRow
  Dim iColPos       As Long
  
  'Define using the AddCol method
  With oRow
    .AddCol "ClientID", 1468&, 8&, 1&
    .AddCol "Name", "John Doe", 0&, 0&
    .AddCol "Address", "47 Main Street", 0&, 0&
    .AddCol "City", "Geneva", 0&, 0&
    .AddCol "State", "Switzerland", 0&, 0&
    'insert the Zip column after the Address column
    iColPos = .ColPos("Address")
    .AddCol "Zip", "12345", 0&, 0&, plInsertAfter:=iColPos
  End With
  ListDump oRow, "AddCol() method"
  Output "Trying to insert a column with an existing name but different case..."
  On Error Resume Next
  oRow.ColCaseSensitive = False  'true by default !
  oRow.AddCol "NaME", "John Doe", 0&, 0&
  If Err.Number <> 0 Then
    If Err.Number = 457& Then
      OutputLn "All right, a duplicate key error has been trapped"
    Else
      If Err.Number = 5 Then  'Array already sorted
        OutputLn "All right, the (custom) error #5 (array already sorted) has been trapped"
      Else
        Test_SetSuccess LOCAL_ERR_CTX, False, "An error occurred, but should have been #5 or better, #457"
        GoTo TestPopulateWithAddCol_Exit
      End If
    End If
  Else
    Test_SetSuccess LOCAL_ERR_CTX, False, "An error should have occurred when inserting a column with the same name"
    GoTo TestPopulateWithAddCol_Exit
  End If

  Output "Changing key case and trying again to insert a column with an existing name but different case..."
  On Error Resume Next
  Err.Clear
  Set oRow = New CRow
  oRow.ColCaseSensitive = True
  'Define using the AddCol method
  With oRow
    .AddCol "ClientID", 1468&, 8&, 1&
    .AddCol "Name", "John Doe", 0&, 0&
    .AddCol "Address", "47 Main Street", 0&, 0&
    .AddCol "City", "Geneva", 0&, 0&
    .AddCol "State", "Switzerland", 0&, 0&
    'insert the Zip column after the Address column
    iColPos = .ColPos("Address")
    .AddCol "Zip", "12345", 0&, 0&, plInsertAfter:=iColPos
  End With
  oRow.AddCol "NaME", "John Doe", 0&, 0&
  If Err.Number = 0 Then
    ListDump oRow, "Case sensitive column names"
  Else
    Test_SetSuccess LOCAL_ERR_CTX, False, "Couldn't add a column with same name different case: " & Err.Description
    GoTo TestPopulateWithAddCol_Exit
  End If
  
TestPopulateWithAddCol_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestPopulateWithAddCol_Err:
  OutputLn LOCAL_ERR_CTX & " error: " & Err.Description
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestPopulateWithAddCol_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestRowArrayDefine()
  Const LOCAL_ERR_CTX As String = "TestRowArrayDefine"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Defining column set Using ""on the fly"" arrays with the VB Array() method"

  Dim oRow          As New CRow
  
  On Error GoTo TestRowArrayDefine_Err
  'ArrayDefine parameters (arrays): name, type, size, flags
  oRow.ArrayDefine Array("ClientID", "Name", "Address", "Zip", "City", "State"), _
                  Array(vbLong, vbString, vbString, vbString, vbString, vbString), _
                  Array(8&, 0&, 0&, 0&, 0&, 0&), _
                  Array(1&, 0&, 0&, 0&, 0&, 0&)
  oRow.Assign 1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland"
  ListDump oRow, "ArrayDefine() method"
  
  Test_SetSuccess LOCAL_ERR_CTX, True
  
TestRowArrayDefine_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestRowArrayDefine_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestRowArrayDefine_Exit
End Sub
  
#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestRowDefine()
  Const LOCAL_ERR_CTX As String = "TestRowDefine"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Defining column set Using the Define method"
  
  On Error GoTo TestRowDefine_Err

  Dim oRow          As New CRow
  'name, type, size, flags
  oRow.Define "ClientID", vbLong, 8&, 1&, _
              "Name", vbString, 0&, 0&, _
              "Address", vbString, 0&, 0&, _
              "Zip", vbString, 0&, 0&, _
              "City", vbString, 0&, 0&, _
              "State", vbString, 0&, 0&
  oRow.Assign 1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland"
  ListDump oRow, "Define() method"

  Test_SetSuccess LOCAL_ERR_CTX, True
  
TestRowDefine_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestRowDefine_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestRowDefine_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestArrayAssign()
  Const LOCAL_ERR_CTX As String = "TestArrayAssign"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Populate column set Using the ArrayAssign method"

  On Error GoTo TestArrayAssign_Err

  Dim oRow          As New CRow
  'name, type, size, flags
  oRow.Define "ClientID", vbLong, 8&, 1&, _
              "Name", vbString, 0&, 0&, _
              "Address", vbString, 0&, 0&, _
              "Zip", vbString, 0&, 0&, _
              "City", vbString, 0&, 0&, _
              "State", vbString, 0&, 0&
  oRow.ArrayAssign Array(1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland")
  ListDump oRow, "Define() method"

  Test_SetSuccess LOCAL_ERR_CTX, True

TestArrayAssign_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestArrayAssign_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestArrayAssign_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestCopyAndCloningRows()
  Const LOCAL_ERR_CTX As String = "TestCopyAndCloningRows"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Copying and cloning rows"

  On Error GoTo TestCopyAndCloningRows_Err

  Dim oRow1     As New CRow
  Dim oRow2     As New CRow
  Dim oRowClone As CRow
  
  oRow1.Define "ClientID", vbLong, 8&, 1&, _
                "Name", vbString, 0&, 0&, _
                "Address", vbString, 0&, 0&, _
                "Zip", vbString, 0&, 0&, _
                "City", vbString, 0&, 0&, _
                "State", vbString, 0&, 0&
  oRow1.Assign 1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland"
  
  'Copy row
  oRow2.CopyFrom oRow1
  ListDump oRow2, "oRow2 copied from oRow1"
  
  'Clone row
  Set oRowClone = oRow1.Clone()
  ListDump oRowClone, "oRowClone created from oRow1"

  'Embedded rows
  Dim rowParent As New CRow
  Dim rowChild  As New CRow
  Dim rowClone  As CRow
  
  rowParent.ArrayDefine Array("Number", "ChildRow", "Text"), Array(vbInteger, vbObject, vbString)
  rowParent("Number") = 1
  rowParent("Text") = "Parent row"
  rowChild.ArrayDefine Array("value1", "value2", "value3"), Array(vbVariant, vbVariant, vbVariant)
  rowChild("value1") = 1
  rowChild("value2") = 2
  rowChild("value3") = 3
  Set rowParent("ChildRow") = rowChild
  
  'Clone parent
  Set rowClone = rowParent.Clone()
  
  'Verify that child in parent and clone is same object
  'The clone method does a SHALLOW copy
  OutputLn "Updating 'value1' in rowParent should be echoed by rowClone (same object ref)"
  rowClone("ChildRow")("value1") = 99
  OutputLn "rowClone('ChildRow')('value1')=" & rowClone("ChildRow")("value1")
  OutputLn "rowParent('ChildRow')('value1')=" & rowParent("ChildRow")("value1")
  Test_Value 99, rowParent("ChildRow")("value1"), "Modifying one ref should update same object"
  
  Test_SetSuccess LOCAL_ERR_CTX, True
  
TestCopyAndCloningRows_Exit:
  Set rowClone = Nothing
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestCopyAndCloningRows_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestCopyAndCloningRows_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestRowErrors()
  Const LOCAL_ERR_CTX As String = "TestRowErrors"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test that some methods fail as expected"
  
  On Error GoTo TestRowErrors_Err

  Dim oRow          As New CRow
  'name, type, size, flags
  oRow.Define "ClientID", vbLong, 8&, 1&, _
              "Name", vbString, 0&, 0&, _
              "Address", vbString, 0&, 0&, _
              "Zip", vbString, 0&, 0&, _
              "City", vbString, 0&, 0&, _
              "State", vbString, 0&, 0&
  oRow.ArrayAssign Array(1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland")
  ListDump oRow, "Define() method"
  
  Dim vTestValue As Variant
  Dim lErr As Long
  
  'Non existing column
  On Error Resume Next
  vTestValue = oRow("Non existing column name")
  lErr = Err.Number
  On Error GoTo TestRowErrors_Err
  Test_Value 5&, lErr, "Unexpected error number"
  Err.Clear
  
  'Invalid numeric index
  On Error Resume Next
  vTestValue = oRow(oRow.ColCount + 1)
  lErr = Err.Number
  On Error GoTo TestRowErrors_Err
  Test_Value 5&, lErr, "Unexpected error number"
  
  'Invalid string numeric index
  On Error Resume Next
  vTestValue = oRow("#" & (oRow.ColCount + 1))
  lErr = Err.Number
  On Error GoTo TestRowErrors_Err
  Test_Value 5&, lErr, "Unexpected error number"

  'Invalid column index on ColName
  On Error Resume Next
  vTestValue = oRow.ColName(oRow.ColCount + 1)
  lErr = Err.Number
  On Error GoTo TestRowErrors_Err
  Test_Value 5&, lErr, "Unexpected error number"
  
  'Invalid column index on ColSize
  On Error Resume Next
  vTestValue = oRow.ColSize("--nope--")
  lErr = Err.Number
  On Error GoTo TestRowErrors_Err
  Test_Value 5&, lErr, "Unexpected error number"
  
  Test_SetSuccess LOCAL_ERR_CTX, True

TestRowErrors_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestRowErrors_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestRowErrors_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestMergeRows()
  Const LOCAL_ERR_CTX As String = "TestMergeRows"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Merging rows"
  
  On Error GoTo TestMergeRows_Err

  Dim oRow1       As New CRow
  Dim oRow2       As New CRow
  Dim iColCount1  As Integer
  Dim iColCount2  As Integer
  
  oRow1.Define "ClientID", vbLong, 8&, 1&, _
                "Name", vbString, 0&, 0&, _
                "Address", vbString, 0&, 0&
  oRow1.Assign 1468&, "John Doe", "47 Main Street"
  iColCount1 = oRow1.ColCount
  ListDump oRow1, "Row1"
  
  oRow2.Define "Zip", vbString, 0&, 0&, _
                "Name", vbString, 0&, 0&, _
                "Address", vbString, 0&, 0&, _
                "City", vbString, 0&, 0&, _
                "State", vbString, 0&, 0&
  oRow2.Assign "12345", "Patrick Doe", "1 nowhere street", "Geneva", "Switzerland"
  iColCount2 = oRow1.ColCount
  ListDump oRow2, "Row2"
  
  'Merge Row2 into Row1
  'The values of oRow2 overwrite the existing oRow1 values.
  oRow1.Merge oRow2
  ListDump oRow1, "oRow1 merged with oRow2"
  
  Test_Value iColCount1 + iColCount2, oRow1.ColCount, "Invalid columns count after merge"
  Test_Value oRow2("Name"), oRow1("Name"), "rows should be equal"
  
  Test_SetSuccess LOCAL_ERR_CTX, True

TestMergeRows_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestMergeRows_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestMergeRows_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestSharpNotation()
  Const LOCAL_ERR_CTX As String = "TestSharpNotation"
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Populate column set of a CRow object and access columns with # notation"
  
  On Error GoTo TestSharpNotation_Err

  Dim oRow          As New CRow
  Dim iColPos       As Long
  Dim sColName      As String
  
  'Define using the AddCol method
  With oRow
    'With a named column, would be: .AddCol "ClientID", 1468&, 8&, 1&
    .AddCol "", 1468&, 8&, 1&
    .AddCol "Address", "47 Main Street", 0&, 0&
    'With a named column, would be: .AddCol "Name", "John Doe", 0&, 0&, plInsertAfter:=1&
    .AddCol "", "John Doe", 0&, 0&, plInsertAfter:=1&
    .AddCol "State", "Switzerland", 0&, 0&
    .AddCol "City", "Geneva", 0&, 0&, plInsertAfter:=3&
    'insert the Zip column after the Address column
    iColPos = .ColPos("Address")
    .AddCol "Zip", "12345", 0&, 0&, plInsertAfter:=iColPos
    
    ListDump oRow, "Test row"
    Test_Value 6, .ColCount, "Unexpected number of columns in test row"
    
    For iColPos = 1 To .ColCount
      sColName = .ColName(iColPos)
      If Len(sColName) > 0 Then
        Output "ColValue('#" & iColPos & "')=" & .ColValue("#" & iColPos) & " == ColValue('" & sColName & "') ? "
        If .ColValue("#" & iColPos) = .ColValue(sColName) Then
          OutputLn "OK"
        Else
          Test_SetSuccess LOCAL_ERR_CTX, False, "Values do not match for ColPos=" & iColPos & " and ColName='" & sColName & "'"
          GoTo TestSharpNotation_Exit
        End If
      Else
        OutputLn "Column #" & iColPos & " has no name. ColValue('#" & iColPos & "')=" & .ColValue("#" & iColPos)
        If iColPos = 1 Then
          Test_Value 1468&, .ColValue("#" & iColPos), "Expected value in unnamed column 'ClientID' not found"
        Else
          Test_Value "John Doe", .ColValue("#" & iColPos), "Expected value in unnamed column 'Name' not found"
        End If
      End If
    Next iColPos
  End With
  
  Test_SetSuccess LOCAL_ERR_CTX, True

TestSharpNotation_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestSharpNotation_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, LOCAL_ERR_CTX & " FAILED: " & Err.Description
  Resume TestSharpNotation_Exit
  Resume
End Sub

#If TWINBASIC Then
End Module
#End If
