Attribute VB_Name = "MJsonTests"
#If TWINBASIC Then
[ TestFixture ]
Private Module MJsonTests
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Public Sub Test_RunAllTests()
  On Error GoTo Test_RunAllTests_Err
  
  Test_BeginSuite "MJsonTests"

  TestRow2Json
  TestDatabase_Orders
  TestJsonReadAndOutput
  TestJsonInputFiles
  TestDatabase_Queries
  
Test_RunAllTests_Exit:
  Test_EndSuite
  Exit Sub

Test_RunAllTests_Err:
  ConOutLn "Test_RunAllTests() failed: " & Err.Description
  Resume Test_RunAllTests_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestRow2Json()
  Const LOCAL_ERR_CTX As String = "TestRow2Json"
  On Error GoTo TestRow2Json_Err
  Test_BeginTestProc LOCAL_ERR_CTX
  
  Dim oRow    As New CRow

  oRow.ArrayDefine Array("LastModified", "FirstName", "BirthDate", "Amount", "Notes"), _
                   Array(vbVariant, vbString, vbVariant, vbCurrency, vbString)
  
  oRow.Assign #1/28/2020 1:30:00 PM#, "John", #12/1/1970#, 18914.9, "Notes text"

  Dim oConv     As New CJsonConverter
  Dim sJson     As String
  Dim sExpected As String
  
  sExpected = "{""LastModified"":""2020-01-28T12:30:00.000Z"",""FirstName"":""John"",""BirthDate"":""1970-11-30T23:00:00.000Z"",""Amount"":18914.9,""Notes"":""Notes text""}"
  sJson = oConv.ConvertToJson(oRow)
  Test_Value sExpected, sJson, "Test json conversion success"
  Test_SetSuccess LOCAL_ERR_CTX, True
  
TestRow2Json_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestRow2Json_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestRow2Json_Exit
End Sub

Private Function GetTestDatabaseFile() As String
  Dim sFilename As String
  sFilename = Test_GetParam(TESTPARAM_NWINDPATH) & ""
  If Len(sFilename) = 0 Then
    sFilename = CombinePath(GetTestDataInputDirectory(), "Ms Access Northwind Database.accdb")
  End If
  GetTestDatabaseFile = sFilename
End Function

Public Sub TestDatabase_QueryOrders()
  Const LOCAL_ERR_CTX As String = "TestDatabase_Orders"
  On Error GoTo TestDatabase_Orders_Err
  Test_BeginTestProc LOCAL_ERR_CTX
  
  Dim fOK         As Boolean
  Dim lstOrders   As CList
  Dim cnData      As ADODB.Connection
  Dim sDbFilename As String
  Dim SQL         As String
  
  Dim sTestDatabase     As String
  
  sTestDatabase = GetTestDatabaseFile()
  If Not ExistFile(sTestDatabase) Then
    'test dabase not found
    Test_SetSuccess LOCAL_ERR_CTX, False, "Missing test database: " & sTestDatabase
    GoTo TestDatabase_Orders_Exit
  End If

  Set lstOrders = New CList
  Set cnData = ADOOpenConnection(ADOGetAccessConnString(sTestDatabase), "")
  
  SQL = "SELECT * FROM [Order Summary]"
  fOK = ADOGetSnapshotList(cnData, SQL, lstOrders)
  If Not fOK Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to query source database [" & sTestDatabase & "] error #" & ADOLastErr() & ": " & ADOLastErrDesc()
    GoTo TestDatabase_Orders_Exit
  End If
  cnData.Close
  
  Test_SetSuccess LOCAL_ERR_CTX, True

TestDatabase_Orders_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  On Error Resume Next
  Set cnData = Nothing
  Set lstOrders = Nothing
  Exit Sub
  
TestDatabase_Orders_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestDatabase_Orders_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestDatabase_Orders()
  Const LOCAL_ERR_CTX As String = "TestDatabase_Orders"
  On Error GoTo TestDatabase_Orders_Err
  Test_BeginTestProc LOCAL_ERR_CTX
  
  Dim fOK         As Boolean
  Dim lstOrders   As CList
  Dim cnData      As ADODB.Connection
  Dim sDbFilename As String
  Dim SQL         As String
  
  Dim sTestDatabase     As String
  
  sTestDatabase = GetTestDatabaseFile()
  If Not ExistFile(sTestDatabase) Then
    'test dabase not found
    Test_SetSuccess LOCAL_ERR_CTX, False, "Missing test database: " & sTestDatabase
    GoTo TestDatabase_Orders_Exit
  End If

  Set lstOrders = New CList
  Set cnData = ADOOpenConnection(ADOGetAccessConnString(sTestDatabase), "")
  
  SQL = "SELECT * FROM [Order Summary]"
  fOK = ADOGetSnapshotList(cnData, SQL, lstOrders)
  If Not fOK Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to query source database [" & sTestDatabase & "] error #" & ADOLastErr() & ": " & ADOLastErrDesc()
    GoTo TestDatabase_Orders_Exit
  End If
  cnData.Close
  
  Dim oConv   As New CJsonConverter
  Dim sJson   As String
  Dim sTest   As String
  
  sJson = oConv.ConvertToJson(lstOrders)
  sTest = """Ship Address"":""789 27th Street"",""Paid Date"":""2006-01-14T23:00:00.000Z"",""Status"":""Closed""}]"
  Test_Value sTest, Right$(sJson, Len(sTest)), "Expecting valid JSON conversion from orders CList"
  
  Test_SetSuccess LOCAL_ERR_CTX, True

TestDatabase_Orders_Exit:
  Test_SetSuccess LOCAL_ERR_CTX, True '/**/ force always OK
  Test_EndTestProc LOCAL_ERR_CTX
  On Error Resume Next
  Set cnData = Nothing
  Set lstOrders = Nothing
  Exit Sub
  
TestDatabase_Orders_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestDatabase_Orders_Exit
End Sub

Private Sub TryJsonFileToRowList(ByRef poRetJson As Object, ByVal psFilename As String)
  Const LOCAL_ERR_CTX As String = "TryJsonFileToRowList"
  On Error GoTo TryJsonFileToRowList_Err
  
  Dim sJsonFilename   As String
  Dim sJson           As String
  Dim oConv           As New CJsonConverter
  Dim oJson           As Object
  
  Set poRetJson = Nothing
  sJsonFilename = CombinePath(GetTestDataInputDirectory(), psFilename)
  OutputLn "Read JSON file: " & sJsonFilename
  sJson = GetFileText(sJsonFilename, "utf8")
  If Len(sJson) = 0 Then
    OutputLn "File [" & sJsonFilename & "] is empty"
    GoTo TryJsonFileToRowList_Exit
  End If
  
  OutputLn "<JSON>"
  OutputLn sJson
  OutputLn "</JSON>"
  OutputLn "Convert JSON to Object"
  Set oJson = oConv.ParseJson(sJson)
  OutputLn "Type name of parsed object: " & TypeName(oJson)

  Dim sOutputFile As String
  sOutputFile = psFilename & ".listdump.html"
  OpenTraceOutputFile sOutputFile, True
  ListDump oJson, psFilename, "*:15", pfDeepDump:=True, pfHTML:=True, pfIgnoreUnattended:=True
  CloseTraceOutputFile
  ViewTraceOutputFile sOutputFile
  
  Set poRetJson = oJson
  
TryJsonFileToRowList_Exit:
  Set oJson = Nothing
  Exit Sub

TryJsonFileToRowList_Err:
  OutputLn LOCAL_ERR_CTX & " failed, error #" & Err.Number & ": " & Err.Description & ", on file [" & psFilename & "]"
  Resume TryJsonFileToRowList_Exit
End Sub

Public Sub TestJsonInputFiles()
  Dim avPath      As Variant
  Dim oJson       As Object
  Dim vValue      As Variant
  
  'GoTo skip
  
  TryJsonFileToRowList oJson, "test_json_1.txt"

  'A path is a array which each element is either a json element name, or
  'a array of two elements, the first being a json element name, and the
  'second a long integer index in [1..rowcount].
  avPath = Array("glossary", "GlossDiv", "GlossList", "GlossEntry", "GlossDef", "GlossSeeAlso", Array("col1", 2))
  vValue = FollowPath(oJson, avPath)
  OutputLn (vValue)
  
  Test_Value "XML", vValue, "Path unreachable"

  TryJsonFileToRowList oJson, "test_json_2.txt"

  TryJsonFileToRowList oJson, "test_json_3.txt"
skip:
  TryJsonFileToRowList oJson, "test_json_4.txt"

End Sub

Private Sub WriteJsonToFile(ByRef psJson As String, ByVal psFilename As String)
  Const LOCAL_ERR_CTX As String = "WriteJsonToFile"
  On Error GoTo WriteJsonToFile_Err
  
  Dim hFile       As String
  Dim fIsOpen     As Boolean
  Dim lJsonLen    As Long
  Dim iAscW       As Integer
  Dim i           As Long
  
  If ExistFile(psFilename) Then
    Kill psFilename
  End If
  
  hFile = FreeFile
  Open psFilename For Binary As #hFile
  fIsOpen = True
  
  lJsonLen = Len(psJson)
  If lJsonLen > 0 Then
    For i = 1 To lJsonLen
      iAscW = AscW(Mid$(psJson, i, 1))
      Put #hFile, , iAscW
    Next i
  End If
  
WriteJsonToFile_Exit:
  If fIsOpen Then
    Close hFile
  End If
  Exit Sub

WriteJsonToFile_Err:
  OutputLn LOCAL_ERR_CTX & " failed, error #" & Err.Number & ": " & Err.Description & ", on file [" & psFilename & "]"
  Resume WriteJsonToFile_Exit
End Sub

Public Sub ViewFile(ByVal psFilename As String)
  If Nz(Test_GetParam(SUITEPARAM_FAKEVIEWFILE), False) = False Then
    If ExistFile(psFilename) Then
      Shell TraceEditorEXE() & " " & psFilename, vbMaximizedFocus
    Else
      MsgBox "Cannot open file [" & psFilename & "] doesn't exist", vbCritical
    End If
  End If
End Sub

Public Sub TestJsonReadAndOutput()
  Dim oJson       As Object
  Dim oConv       As New CJsonConverter
  Dim sJson       As String
  Dim sFilename   As String
  Dim i           As Integer
  For i = 1 To 4
    sFilename = CombinePath(GetTestDataOutputDirectory(), "test_json_" & i & ".parsed.txt")
    TryJsonFileToRowList oJson, "test_json_" & i & ".txt"
    sJson = oConv.ConvertToJson(oJson)
    WriteJsonToFile sJson, sFilename
    ViewFile sFilename
  Next i
End Sub

Public Function JsonTrim(ByVal psText As String) As String
  Dim i As Long
  Dim sChar As String
  
  'if psText has any other char than ["{","}","[","]"],
  'except for whitespace [" ",vbTab,vbCr,vbLf], then
  'we return an empty string, otherwise we return
  'the original string.
  
  For i = 1 To Len(psText)
    sChar = Mid$(psText, i, 1)
    If (sChar <> "{") And (sChar <> "[") And (sChar <> "]") And (sChar <> "}") Then
      If (sChar <> " ") And (sChar <> vbTab) And (sChar <> vbCr) And (sChar <> vbLf) Then
        JsonTrim = psText
        Exit Function
      End If
    End If
  Next i
End Function

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestDatabase_Queries()
  Const LOCAL_ERR_CTX As String = "TestDatabase_Queries"
  On Error GoTo TestDatabase_Queries_Err
  Test_BeginTestProc LOCAL_ERR_CTX
  
  Dim fOK         As Boolean
  Dim lstData   As CList
  Dim cnData      As ADODB.Connection
  Dim sDbFilename As String
  Dim avSQL       As Variant
  Dim iCount      As Integer
  Dim i           As Integer
  
  Dim sTestDatabase     As String
  
  sTestDatabase = GetTestDatabaseFile()
  If Not ExistFile(sTestDatabase) Then
    'test dabase not found
    Test_SetSuccess LOCAL_ERR_CTX, False, "Missing test database: " & sTestDatabase
    GoTo TestDatabase_Queries_Exit
  End If

  Set lstData = New CList
  Set cnData = ADOOpenConnection(ADOGetAccessConnString(sTestDatabase), "")
  
  avSQL = Array( _
      "SELECT * FROM [Order Summary]", _
      "SELECT * FROM [Order Details Extended]", _
      "SELECT * FROM [Order Price Totals]", _
      "SELECT * FROM [Product Category Sales by Date]", _
      "SELECT * FROM [Product Sales by Category]", _
      "SELECT * FROM [Product Sales Qty by Employee and Date]", _
      "SELECT * FROM [Shippers Extended]", _
      "SELECT * FROM [Sales Analysis]" _
    )
  
  Dim oConv   As New CJsonConverter
  Dim sJson   As String
  Dim sTest   As String
  Dim sQuery  As String
  Dim k       As Integer
  
  iCount = UBound(avSQL) - LBound(avSQL) + 1&
  For i = 0 To iCount - 1
    sQuery = avSQL(i)
    Test_Comment "Running query (" & (i + 1) & "/" & iCount & ") [" & sQuery & "]"
    fOK = ADOGetSnapshotList(cnData, sQuery, lstData)
    If Not fOK Then
      Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to query source database [" & sTestDatabase & "] error #" & ADOLastErr() & ": " & ADOLastErrDesc()
      GoTo TestDatabase_Queries_Exit
    End If
    
    sJson = oConv.ConvertToJson(lstData)
    sTest = JsonTrim(sJson)
    Test_ValueNotEqual "", sTest, "Checking JSON conversion <json>" & Left$(sTest, 80) & IIf(Len(sTest) > 80, " ...", "</json>")
  Next i
  cnData.Close
  
  Test_SetSuccess LOCAL_ERR_CTX, True

TestDatabase_Queries_Exit:
  On Error Resume Next
  If Not cnData Is Nothing Then
    cnData.Close
  End If
  Test_EndTestProc LOCAL_ERR_CTX
  On Error Resume Next
  Set cnData = Nothing
  Set lstData = Nothing
  Exit Sub
  
TestDatabase_Queries_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestDatabase_Queries_Exit
End Sub

#If TWINBASIC Then
End Module
#End If


