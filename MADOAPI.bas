Attribute VB_Name = "MADOAPI"
' Compile flag    ¦ Comment
'-----------------+--------------------------------------------------------------------
' MSXML           ¦ To include XML features (see ADOGetSnapshotXML())
' MSACCESS        ¦ Activate VBA (versus VB) language features
' ROWLISTLIB      ¦ Compile functions using advanced CRow and CList classes.
' CDBCONNECTION   ¦ Activate CDBConnection class coupling
' (Created 09/06/2007)
#If TWINBASIC Then
Module MADOAPI
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

#If ROWLISTLIB Then
'Special ADO column flags
Public Const ADOCOL_AUTOINCREMENT   As Long = &H1000&
Public Const ADOCOL_ISNULLABLE      As Long = &H2000&
Public Const ADOCOL_LONGTEXT        As Long = &H4000&
Public Const ADOCOL_BLOB            As Long = &H8000&
#End If

'Note 1
'------
'In 64 bits the record count is a long long, which we'll not handle
'has we don't care here; but we raise an error for that.
'There are multiple solutions to that, a possible one would be to
'store the LongLong in a module variable and expose it thru
'a new api function.
Public Const ADOERR_TOOLARGE        As Long = -1
Public Const ADOERRSTRING_TOOLARGE  As String = "The returned record count is too large to be handled by this application" 'no need to localise that

Public Enum eAdoDatabaseEngine
  edbEngineAccess = 0     'default database engine
  edbEngineSQLServer = 1
End Enum

Private meDbEngine As eAdoDatabaseEngine

Private Const QUOTE_CHAR      As String = "'"

'Transaction stacking V01.01.00
Private mlStackedTransCt        As Long
Private mfRequireTransFeature   As Boolean  'Not that we inverted meaning from CDBConnection.mfIgnoreTransFeature

'Error context
Private mlErr           As Long
Private msErr           As String
Private msErrCtx        As String
Private msLastSQLState  As String

Private Sub ClearErr()
  mlErr = 0&
  msErr = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErr As Long, ByVal psErr As String)
  msErrCtx = psErrCtx
  mlErr = plErr
  msErr = psErr
End Sub

Public Function ADOLastErr() As Long
  ADOLastErr = mlErr
End Function

Public Function ADOLastErrDesc() As String
  ADOLastErrDesc = msErr
End Function

Public Function ADOLastErrCtx() As String
  ADOLastErrCtx = msErrCtx
End Function

Public Function ADOLastSQLState() As String
  ADOLastSQLState = msLastSQLState
End Function

Public Function ADOGetAccessConnString(ByVal psDatabasePathname As String) As String
  Dim sConnString   As String
#If Win64 Then
  sConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & psDatabasePathname
#Else
  If GetFileExt(psDatabasePathname) = "accdb" Then
    sConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & psDatabasePathname
  Else
    sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & psDatabasePathname
  End If
#End If
  ADOGetAccessConnString = sConnString
End Function

'
' Connect to an ADO database
'

Public Function ADOGetEngine() As eAdoDatabaseEngine
  ADOGetEngine = meDbEngine
End Function

Public Sub ADOSetEngine(ByVal peEngine As eAdoDatabaseEngine)
  meDbEngine = peEngine
End Sub

Public Function ADOOpenConnection(ByRef psConnString As String, ByRef psDatabase As String) As ADODB.Connection
  Dim cnNewConn       As ADODB.Connection
  Dim vErr            As Variant
  Dim sErrInfo        As String
  
  On Error GoTo ADOOpenConnection_Err
  ClearErr
  sErrInfo = "ADOOpenConnection('" & psConnString & "')"
  
  Set cnNewConn = New ADODB.Connection
  cnNewConn.ConnectionString = psConnString
  If Len(psDatabase) Then
    cnNewConn.Open psDatabase
  Else
    cnNewConn.Open
  End If

  Set ADOOpenConnection = cnNewConn
  Exit Function
ADOOpenConnection_Err:
  SetErr "ADOOpenConnection", Err.Number, Err.Description & " (" & sErrInfo & ")"
  If Not cnNewConn Is Nothing Then
    If cnNewConn.Errors.Count Then
      On Error Resume Next
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In cnNewConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
        msLastSQLState = vErr.SQLState
      Next
    End If
  End If
  Set cnNewConn = Nothing
End Function

'
' Quickly get the snapshot of a query.
'
Public Function ADOGetSnapshotData(pcnConn As ADODB.Connection, ByRef psSQL As String, ByRef pvRetData As Variant, Optional ByVal plReadMax As Long = 0&) As Long
  Dim rsSnap    As New ADODB.Recordset
  Dim lRows     As Long
  Dim vErr      As Variant
  Dim lCount    As Long
  
  On Error GoTo ADOGetSnapshotData_Err
  ClearErr
  
  pvRetData = Null
  rsSnap.Open psSQL, pcnConn, adOpenStatic, adLockReadOnly, adCmdText
  If Not rsSnap.EOF Then
    rsSnap.MoveLast
    rsSnap.MoveFirst
    #If Win64 Then
      Const MAX_LONG = 2147483647  'Long (32bits) [-2,147,483,648 and 2,147,483,647]
      If rsSnap.RecordCount > MAX_LONG Then
        SetErr "ADOGetSnapshotData", ADOERR_TOOLARGE, ADOERRSTRING_TOOLARGE
        rsSnap.Close
        Exit Function
      End If
    #End If
    lRows = CLng(rsSnap.RecordCount)
    If plReadMax = 0& Then
      pvRetData = rsSnap.GetRows()
      ADOGetSnapshotData = lRows
    Else
      pvRetData = rsSnap.GetRows(plReadMax)
      ADOGetSnapshotData = Min(lRows, plReadMax)
    End If
  End If
  rsSnap.Close
  
  Exit Function

ADOGetSnapshotData_Err:
  SetErr "ADOGetSnapshotData", Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
End Function

Public Function ADOSnapshot(ByVal pcnConn As ADODB.Connection, ByRef psSQL As String) As ADODB.Recordset
  Dim rsSnap    As New ADODB.Recordset
  Dim lRows     As Long
  Dim vErr      As Variant
  
  On Error GoTo ADOSnapshot_Err
  ClearErr
  
  rsSnap.Open psSQL, pcnConn, adOpenStatic, adLockReadOnly, adCmdText
  If Not rsSnap.EOF Then
    rsSnap.MoveLast
    rsSnap.MoveFirst
  End If
  
  Set ADOSnapshot = rsSnap
  Exit Function

ADOSnapshot_Err:
  SetErr "ADOSnapshot", Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Set pcnConn = Nothing
End Function

'
' Get the snapshot of a query as XML.
'
#If MSXML Then

'source: adapted from a non retrieved source / internet

Public Function ADOGetSnapshotXML(pcnConn As ADODB.Connection, ByRef psSQL As String, ByVal psGroupName As String, ByVal psEntryName As String, ByRef pvRetData As Variant, ByRef poRetDOMdoc As DOMDocument60) As Long
  Dim rsSnap    As New ADODB.Recordset
  Dim lRows     As Long
  Dim vErr      As Variant
  Dim oDOMdoc1  As DOMDocument60
  
  On Error GoTo ADOGetSnapshotXML_Err
  ClearErr
  
  pvRetData = Null
  Set poRetDOMdoc = Nothing
  rsSnap.Open psSQL, pcnConn, adOpenStatic, adLockReadOnly, adCmdText
  If Not rsSnap.EOF Then
    rsSnap.MoveLast
    rsSnap.MoveFirst
    lRows = rsSnap.RecordCount
    pvRetData = rsSnap.GetRows()
  End If
  Set oDOMdoc1 = New DOMDocument60
  rsSnap.Save oDOMdoc1, adPersistXML
  'debug: oDOMdoc1.Save NormalizePath(StripFileName(CurrentDb.Name)) & "temp.xml"
  Set poRetDOMdoc = ADOConvertToElementTree(oDOMdoc1, psGroupName, psEntryName)
  rsSnap.Close
  
  Set oDOMdoc1 = Nothing
  ADOGetSnapshotXML = lRows
  Exit Function

ADOGetSnapshotXML_Err:
  SetErr "ADOGetSnapshotXML", Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Set pcnConn = Nothing
  Set oDOMdoc1 = Nothing
End Function

Public Function ADOConvertToElementTree(rdxml As DOMDocument60, GroupName As String, entryName As String) As DOMDocument60
  Dim xmlDoc      As DOMDocument60
  Dim groupNode   As IXMLDOMElement
  Dim rowNode     As IXMLDOMElement
  Dim elemNode    As IXMLDOMElement
  Dim itemNode    As IXMLDOMElement
  Dim attrNode    As IXMLDOMNode
  Dim rootNode    As IXMLDOMNode
  
  Set xmlDoc = New DOMDocument60

  If Len(GroupName) Then
    Set rootNode = xmlDoc.createElement(GroupName)
    Set xmlDoc.DocumentElement = rootNode
  Else
    Set rootNode = xmlDoc.DocumentElement
  End If
  
  rdxml.SetProperty "SelectionLanguage", "XPath"
  rdxml.SetProperty "SelectionNamespaces", "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' " & _
          "xmlns:sp='http://schemas.microsoft.com/sharepoint/soap/' " & _
          "xmlns:rs='urn:schemas-microsoft-com:rowset' " & _
          "xmlns:z='#RowsetSchema'"
  
  'For Each rowNode In rdxml.SelectNodes("//rs:data/z:row")
  For Each rowNode In rdxml.SelectNodes("//rs:data/z:row")
    If Not rootNode Is Nothing Then
      Set itemNode = rootNode.appendChild( _
             xmlDoc.createElement(entryName))
    Else
      Set itemNode = xmlDoc.createElement(entryName)
    End If
    For Each attrNode In rowNode.SelectNodes("@*")
      Set elemNode = itemNode.appendChild( _
         xmlDoc.createElement(attrNode.nodeName))
      elemNode.nodeTypedValue = attrNode.nodeTypedValue
    Next
  Next
  Set ADOConvertToElementTree = xmlDoc
End Function

#End If

Public Function ADOLastInsertID(pcnConn As ADODB.Connection) As Variant
  Dim iDummy  As Long
  Dim vData   As Variant
  Select Case meDbEngine
  Case eAdoDatabaseEngine.edbEngineAccess, eAdoDatabaseEngine.edbEngineSQLServer
    iDummy = ADOGetSnapshotData(pcnConn, "SELECT @@IDENTITY", vData)
  Case Else
    'implement more engines here
  End Select
  If iDummy Then
    ADOLastInsertID = vData(0&, 0&)
  Else
    ADOLastInsertID = Null
  End If
End Function

Public Function ADODoQuotes(ByVal sData As String) As String
  If Len(sData) Then
    ADODoQuotes = DoQuotes(sData)
  Else
    ADODoQuotes = "''"
  End If
End Function

'ADOSQLDate... functions use the configured db engine
Public Function ADOSQLDateTime(ByVal dtDate As Date) As String
  If meDbEngine = eAdoDatabaseEngine.edbEngineAccess Then
    ADOSQLDateTime = AccessSQLDateTime(dtDate)
  Else
    ADOSQLDateTime = SQLServerSQLDateTime(dtDate)
  End If
End Function

Public Function ADOSQLDate(ByVal dtDate As Date) As String
  If meDbEngine = eAdoDatabaseEngine.edbEngineAccess Then
    ADOSQLDate = AccessSQLDate(dtDate)
  Else
    ADOSQLDate = SQLServerSQLDate(dtDate)
  End If
End Function

'Specific dialects date/time functions

'Note: no need to use SET DATEFORMAT as general YYYMMDD format used.
Public Function SQLServerSQLDateTime(ByVal dtDate As Date) As String
  SQLServerSQLDateTime = "'" & _
          Format$(Year(dtDate), "0000") & _
          Format$(Month(dtDate), "00") & _
          Format$(Day(dtDate), "00") & " " & _
          Format$(Hour(dtDate), "00") & ":" & _
          Format$(Minute(dtDate), "00") & ":" & _
          Format$(Second(dtDate), "00") & "'"
End Function

'Note: use SET DATEFORMAT us_english
Public Function SQLServerSQLDate(ByVal dtDate As Date) As String
  SQLServerSQLDate = "'" & Month(dtDate) & "-" & Day(dtDate) & "-" & Year(dtDate) & "'"
End Function

Public Function ODBCSQLDate(ByVal pDate As Variant) As String
  Dim strRet As String
  If IsDate(pDate) Then
    strRet = "{d '" & Format$(Year(pDate), "0000") & "-" & Format$(Month(pDate), "00") & "-" & Format$(Day(pDate), "00") & "'}"
    ODBCSQLDate = strRet
  End If
End Function

Public Function ODBCSQLDateTime(ByVal theDate As Variant) As String
  Dim strRet As String
  Dim szDate$, szTime$, szDateTime$
  
  On Error Resume Next
  strRet = "{ts '" & Format$(Year(theDate), "0000") & "-" & Format$(Month(theDate), "00") & "-" & Format$(Day(theDate), "00") & " " & Format$(theDate, "hh:mm:ss") & "'}"
  ODBCSQLDateTime = strRet
End Function

'
' Command Execution
'

'Executes an SQL action statement on the specified connection
Public Function ADOExecSQL(pcnExecute As ADODB.Connection, ByRef psSQL As String, Optional ByRef plRetAffectedCt As Long) As Boolean
  If Not IsMissing(plRetAffectedCt) Then
    ADOExecSQL = ExecSQL(pcnExecute, psSQL, plRetAffectedCt)
  Else
    ADOExecSQL = ExecSQL(pcnExecute, psSQL)
  End If
End Function

Public Function ADOTableExistEx(pcnTarget As ADODB.Connection, ByVal psTableName As String) As Boolean
  ADOTableExistEx = TableExist(pcnTarget, psTableName)
End Function

'
' Private Members
'

Private Function TableExist(pcnTarget As ADODB.Connection, ByRef psTableName As String) As Boolean
  'I use late bound and don't care of which ADO library version will be used here,
  'it just has to work if any ADO library is installed.
  Dim oCatalog  As Object 'ADOX.Catalog
  Dim oTable    As Object 'ADOX.Table
  
  On Error GoTo TableExist_Err
  ClearErr
  
  Set oCatalog = CreateObject("ADOX.Catalog") 'New ADOX.Catalog
  oCatalog.ActiveConnection = pcnTarget
  
  On Error Resume Next
  Set oTable = oCatalog.Tables(psTableName)
  TableExist = (Not oTable Is Nothing)
  'fall thru
TableExist_Err:
  Set oTable = Nothing
  Set oCatalog = Nothing
End Function

'Executes an SQL action statement on the specified connection
Private Function ExecSQL(pcnExecute As ADODB.Connection, ByRef psSQL As String, Optional ByRef plRetAffectedCt As Long) As Boolean
  Dim vErr        As Variant
  Dim cmdSQL      As ADODB.Command
  Dim lDummy      As Long
  
  On Error GoTo ExecSQL_Err
  
  Set cmdSQL = New ADODB.Command
  Set cmdSQL.ActiveConnection = pcnExecute
  cmdSQL.CommandType = adCmdText
  cmdSQL.CommandText = psSQL
  If IsMissing(plRetAffectedCt) Then
    cmdSQL.Execute lDummy, , adExecuteNoRecords
  Else
    cmdSQL.Execute plRetAffectedCt
  End If
  Set cmdSQL = Nothing
  
  ExecSQL = True
  
  Exit Function
ExecSQL_Err:
  SetErr "ExecSQL", Err.Number, Err.Description
  If Not pcnExecute Is Nothing Then
    On Error Resume Next
    If pcnExecute.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnExecute.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
        msLastSQLState = vErr.SQLState
      Next
    End If
  End If
  Set cmdSQL = Nothing
End Function

'
' SQL Dialects
'

'Double quotes where needed for stored procedures
Private Function DoQuotes(ByVal sData As String) As String
  If Len(sData) = 0 Then Exit Function
  Dim iLast As Integer
  Dim sPart As String
  
  'Find first single quote
  iLast = InStr(sData, QUOTE_CHAR)
  While iLast
    'Build the string from the left, include two single quotes
    sPart = sPart & Left$(sData, iLast - 1) & QUOTE_CHAR & QUOTE_CHAR
    'Truncate the working string from the right
    sData = Right$(sData, Len(sData) - iLast)
    'Find next single quote in the remainder
    iLast = InStr(sData, QUOTE_CHAR)
  Wend
  'Put any remaining string data on the end
  sData = sPart & sData
  'Return valid string
  DoQuotes = QUOTE_CHAR & sData & QUOTE_CHAR
End Function

Public Function AccessSQLDateTime(ByVal dtDate As Date) As String
  AccessSQLDateTime = "#" & Month(dtDate) & "-" & Day(dtDate) & "-" & Year(dtDate) & " " & _
          Format$(Hour(dtDate), "00") & ":" & _
          Format$(Minute(dtDate), "00") & ":" & _
          Format$(Second(dtDate), "00") & "#"
End Function

Public Function AccessSQLDate(ByVal dtDate As Date) As String
  AccessSQLDate = "#" & Month(dtDate) & "-" & Day(dtDate) & "-" & Year(dtDate) & "#"
End Function

Public Function ADODMin(ByRef pcnConn As ADODB.Connection, ByVal psExpr As String, ByVal psDomain As String, ByVal psCriteria As String) As Variant
  Dim SQL       As String
  Dim vRes      As Variant
  
  On Error Resume Next
  ADODMin = Null
  SQL = "SELECT MIN(" & psExpr & ") FROM " & psDomain
  If Len(psCriteria) Then
    SQL = SQL & " WHERE " & psCriteria
  End If
  If ADOGetSnapshotData(pcnConn, SQL, vRes) Then ADODMin = vRes(0&, 0&)
  If IsEmpty(ADODMin) Then ADODMin = Null
End Function

Public Function ADODMax(ByRef pcnConn As ADODB.Connection, ByVal psExpr As String, ByVal psDomain As String, ByVal psCriteria As String) As Variant
  Dim SQL       As String
  Dim vRes      As Variant
  
  On Error Resume Next
  ADODMax = Null
  SQL = "SELECT MAX(" & psExpr & ") FROM [" & psDomain & "]"
  If Len(psCriteria) Then
    SQL = SQL & " WHERE " & psCriteria
  End If
  If ADOGetSnapshotData(pcnConn, SQL, vRes) Then ADODMax = vRes(0&, 0&)
  If IsEmpty(ADODMax) Then ADODMax = Null
End Function

Public Function ADODCount(ByRef pcnConn As ADODB.Connection, ByVal psExpr As String, ByVal psDomain As String, ByVal psCriteria As String) As Variant
  Dim SQL       As String
  Dim vRes      As Variant
  
  On Error Resume Next
  ADODCount = Null
  SQL = "SELECT COUNT(" & psExpr & ") FROM [" & psDomain & "]"
  If Len(psCriteria) Then
    SQL = SQL & " WHERE " & psCriteria
  End If
  If ADOGetSnapshotData(pcnConn, SQL, vRes) Then ADODCount = vRes(0&, 0&)
  If IsEmpty(ADODCount) Then ADODCount = Null
End Function

Public Function ADODSum(ByRef pcnConn As ADODB.Connection, ByVal psExpr As String, ByVal psDomain As String, ByVal psCriteria As String) As Variant
  Dim SQL       As String
  Dim vRes      As Variant
  
  On Error Resume Next
  ADODSum = Null
  SQL = "SELECT SUM(" & psExpr & ") FROM [" & psDomain & "]"
  If Len(psCriteria) Then
    SQL = SQL & " WHERE " & psCriteria
  End If
  If ADOGetSnapshotData(pcnConn, SQL, vRes) Then ADODSum = vRes(0&, 0&)
  If IsEmpty(ADODSum) Then ADODSum = Null
End Function

Public Function ADODFind(ByRef pcnConn As ADODB.Connection, ByVal psExpr As String, ByVal psDomain As String, ByVal psCriteria As String) As Variant
  Dim SQL       As String
  Dim vRes      As Variant
  On Error Resume Next
  ADODFind = Null
  SQL = "SELECT " & psExpr & " FROM [" & psDomain & "]"
  If Len(psCriteria) Then
    SQL = SQL & " WHERE " & psCriteria
  End If
  If ADOGetSnapshotData(pcnConn, SQL, vRes) Then ADODFind = vRes(0&, 0&)
  If IsEmpty(ADODFind) Then ADODFind = Null
End Function

#If CDBConnection Then
Public Function ADOGetDBConn(ByRef pcnADOConnection As ADODB.Connection) As CDBConnection
  Dim oNewConn    As CDBConnection
  Set oNewConn = New CDBConnection
  Set oNewConn.Connection = pcnADOConnection
  oNewConn.DontClose = True
  Set ADOGetDBConn = oNewConn
  Set oNewConn = Nothing
End Function
#End If

'Generic method from developpez.net wesite
Public Function ADO_GenericOpenRecordset(ByVal strSQL As String, _
                            Optional ByVal eCursorType As ADODB.CursorTypeEnum = adOpenForwardOnly, _
                            Optional ByVal eLockType As ADODB.LockTypeEnum = adLockReadOnly, _
                            Optional ByVal eCommandType As ADODB.CommandTypeEnum = adCmdUnknown, _
                            Optional oConn As ADODB.Connection, _
                            Optional ByVal bOptimizeEval As Boolean = True, _
                            Optional oCollParamValues As Collection = Nothing) As ADODB.Recordset
  Dim p_oConn As ADODB.Connection
  Dim oCmd As ADODB.Command
  Dim oParam As ADODB.Parameter
  ', oQD As DAO.QueryDef, oParam As DAO.Parameter
  Dim oRS As ADODB.Recordset
  Dim oCollEval As Collection
  Dim sExpr As String, sExprColl As String, vValue As Variant
  Dim i As Integer, v As Variant, bAddItem As Boolean
 
    #If MSACCESS Then
    If oConn Is Nothing Then
        Set p_oConn = CurrentProject.Connection
    End If
    #End If
    Set p_oConn = oConn
 
    If bOptimizeEval Then
        If oCollParamValues Is Nothing Then
            Set oCollEval = New Collection
        Else
            Set oCollEval = oCollParamValues
        End If
    End If
 
    On Error Resume Next
 
    Set oCmd = New ADODB.Command
    Set oCmd.ActiveConnection = p_oConn
    oCmd.CommandText = strSQL
 
    If eCommandType = adCmdUnknown Then
        If Trim(strSQL) Like "SELECT *" Then
            eCommandType = adCmdText
        Else
            eCommandType = adCmdTable
        End If
    End If
    oCmd.CommandType = eCommandType
 
    oCmd.Parameters.Refresh
 
    For Each oParam In oCmd.Parameters
 
        bAddItem = True
        sExprColl = oParam.Name
 
        For i = 1 To 3
 
            Select Case i
            Case 1
                ' 1er passage: prendre "l'expression paramètre", telle quelle
                sExpr = sExprColl
            Case 2
                ' 2ème passage: normaliser "l'expression paramètre"
                sExpr = vbNullString
 
                For Each v In Array("[Forms]!", "[Formulaires]!", "Formulaires!")
                    If InStr(1, sExprColl, v) = 1 Then
                        sExpr = Replace(sExprColl, v, "Forms!")
                        Exit For
                    End If
                Next v
 
            Case Else
                ' 3ème et dernier passage: Paramètre non évaluable
                ' Demander la saisie de la valeur du paramètre dans une InputBox
                ' et sortir
                vValue = Null
                vValue = InputBox(sExprColl)
 
                Exit For
            End Select
 
            ' Rechercher l'expression dans la collection
            If bOptimizeEval And Len(sExpr) > 0 Then
                Err.Clear
 
                ' couple <sExpr, Value> déja mémorisé dans la collection ?
                vValue = oCollEval.Item(sExpr)
 
                If Err.Number = 0 Then
                    ' OK - valeur trouvée dans la collection
                    ' sortir de la boucle For
                    bAddItem = False
                    Exit For
                End If
            End If
 
            ' Evaluer l'expression
            Err.Clear
            #If MSACCESS Then
              vValue = Eval(sExpr)
            #Else
              vValue = sExpr
            #End If
 
            Select Case Err.Number
            Case 0
                ' évaluation réussie !
                Exit For
 
            Case 2482, 2451, 2450, 2434, 2425
                ' 2482 = Impossible de touver un nom entré dans l'expression
                ' 2451 = Le nom entré dans l'expression fait référence à un état qui n'existe pas
                ' 2450 = Le nom entré dans l'expression fait référence à un formulaire qui n'existe pas
                ' 2434 = La syntaxe de l'expresion n'est pas correcte
                ' 2425 = L'expression comporte un nom de fonction introuvable
 
            Case Else
                ' autres erreurs ?
            End Select
 
        Next i
 
        If bOptimizeEval And bAddItem Then
            oCollEval.Add vValue, sExprColl
            If Len(sExpr) > 0 And sExpr <> sExprColl Then
                oCollEval.Add vValue, sExpr
            End If
        End If
 
        oParam.Value = vValue
    Next
 
    On Error GoTo 0
 
    Set oRS = New ADODB.Recordset
    oRS.Open oCmd, CursorType:=eCursorType, LockType:=eLockType
 
    Set ADO_GenericOpenRecordset = oRS
 
    Set oParam = Nothing
    Set oRS = Nothing
    Set oCmd = Nothing
    Set p_oConn = Nothing
 
End Function

'
' RowList library dependent methods
'

#If ROWLISTLIB Then

'For the query parameters list, column 1 = name, column 2 = value
Public Function ADOGetSnapshotList(ByRef pcnDatabase As ADODB.Connection, _
                                   ByRef psSQL As String, _
                                   ByRef poRetList As CList, _
                                   Optional ByRef plstParams As CList) As Boolean
  Dim rsSnap          As ADODB.Recordset
  Dim vErr            As Variant
  
  On Error GoTo ADOGetSnapshotList_Err
  ClearErr
  
  If IsMissing(plstParams) Or (plstParams Is Nothing) Then
    Set rsSnap = New ADODB.Recordset
    rsSnap.Open psSQL, pcnDatabase, adOpenStatic, adLockReadOnly, adCmdText
  Else
    Dim i As Long
    Dim oCollParamValues As New Collection
    For i = 1 To plstParams.Count
      'column 1 = name, column 2 = value
      oCollParamValues.Add plstParams(2, i), plstParams(1, i)
    Next i
    Set rsSnap = ADO_GenericOpenRecordset(psSQL, adOpenStatic, adLockReadOnly, adCmdText, pcnDatabase, True, oCollParamValues)
  End If
  
  DefineListFromSet poRetList, rsSnap
  If Not rsSnap.EOF Then
    poRetList.DataArray = rsSnap.GetRows()
    poRetList.SyncWithDataArray
  End If
  rsSnap.Close
  Set rsSnap = Nothing
  ADOGetSnapshotList = True
  Exit Function

ADOGetSnapshotList_Err:
  SetErr "ADOGetSnapshotList", Err.Number, Err.Description
  If Not pcnDatabase Is Nothing Then
    If pcnDatabase.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnDatabase.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Set rsSnap = Nothing
End Function

Public Function ADOGetSnapshotRow(ByRef pcnDatabase As ADODB.Connection, ByRef psSQL As String, ByRef poRetRow As CRow) As Boolean
  Dim oList           As CList
  Dim vErr            As Variant
  
  On Error GoTo ADOGetSnapshotRow_Err
  ClearErr
  ClearRow poRetRow
  Set oList = New CList
  If ADOGetSnapshotList(pcnDatabase, psSQL, oList) Then
    If oList.Count Then
      oList.GetRow poRetRow, 1&
    Else
      oList.DefineRow poRetRow
    End If
    ADOGetSnapshotRow = True
  End If
  poRetRow.Dirty = False
  Set oList = Nothing
  Exit Function
ADOGetSnapshotRow_Err:
  SetErr "ADOGetSnapshotRow", Err.Number, Err.Description
  If Not pcnDatabase Is Nothing Then
    If pcnDatabase.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnDatabase.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Set oList = Nothing
End Function

Public Sub ClearRow(oRow As CRow)
  Dim iColCt        As Long
  Dim iCol          As Long
  Dim vValue        As Variant
  Dim lFlags        As Long
  
  iColCt = oRow.ColCount
  For iCol = 1 To iColCt
    lFlags = oRow.ColFlags(iCol)
    If (lFlags And ADOCOL_ISNULLABLE) = 0& Then
      Select Case oRow.ColType(iCol)
      Case VbVarType.vbBoolean
        vValue = False
      Case VbVarType.vbByte
        vValue = CByte(0)
      Case VbVarType.vbCurrency
        vValue = CCur(0)
      Case VbVarType.vbDecimal
        vValue = CDec(0)
      Case VbVarType.vbDouble
        vValue = CDbl(0)
      Case VbVarType.vbInteger
        vValue = CInt(0)
      Case VbVarType.vbLong
        vValue = 0&
      Case VbVarType.vbSingle
        vValue = CSng(0)
      Case VbVarType.vbDate
        vValue = Now
      Case VbVarType.vbNull
        vValue = Null
      Case VbVarType.vbString
        vValue = ""
      Case Else
        vValue = ""
      End Select
    Else
      vValue = Null
    End If
    oRow(iCol) = vValue
  Next iCol
  oRow.Dirty = False
End Sub

Private Function DefineListFromSet(ByRef oList As CList, ByRef rsTemplate As ADODB.Recordset) As Boolean
  Const LOCAL_ERR_CTX As String = "DefineListFromSet"
  On Error GoTo DefineListFromSet_Err
  ClearErr
  
  Dim i           As Integer
  Dim j           As Integer
  Dim lFlags      As Long
  Dim iFieldCt    As Integer
  Dim avColName   As Variant
  Dim avColType   As Variant
  Dim avColSize   As Variant
  Dim avColFlags  As Variant
  Dim oField      As ADODB.Field
  
  iFieldCt = rsTemplate.Fields.Count
  If iFieldCt = 0 Then Exit Function
  ReDim avColName(1 To iFieldCt)
  ReDim avColType(1 To iFieldCt)
  ReDim avColSize(1 To iFieldCt)
  ReDim avColFlags(1 To iFieldCt)
  j = 1
  For Each oField In rsTemplate.Fields
    lFlags = 0&
    avColName(j) = oField.Name
    avColType(j) = ColTypeToVB(oField.Type, lFlags)
    avColSize(j) = oField.DefinedSize
    If GetFieldProperty(oField, "ISAUTOINCREMENT") Then
      lFlags = lFlags Or ADOCOL_AUTOINCREMENT
    End If
    If oField.Attributes And adFldIsNullable Then
      lFlags = lFlags Or ADOCOL_ISNULLABLE
    End If
    avColFlags(j) = lFlags
    j = j + 1
  Next
  oList.ArrayDefine avColName, avColType, avColSize, avColFlags
  
  DefineListFromSet = True
  
DefineListFromSet_Exit:
  Exit Function
  
DefineListFromSet_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume DefineListFromSet_Exit
End Function

Private Function DefineRowFromSet(poRow As CRow, rsTemplate As ADODB.Recordset) As Boolean
  Const LOCAL_ERR_CTX As String = "DefineRowFromSet"
  On Error GoTo DefineRowFromSet_Err
  ClearErr
  
  Dim i           As Integer
  Dim j           As Integer
  Dim lFlags      As Long
  Dim iFieldCt    As Integer
  Dim avColName   As Variant
  Dim avColType   As Variant
  Dim avColSize   As Variant
  Dim avColFlags  As Variant
  Dim oField      As ADODB.Field
  
  iFieldCt = rsTemplate.Fields.Count
  If iFieldCt = 0 Then Exit Function
  ReDim avColName(1 To iFieldCt)
  ReDim avColType(1 To iFieldCt)
  ReDim avColSize(1 To iFieldCt)
  ReDim avColFlags(1 To iFieldCt)
  j = 1
  For Each oField In rsTemplate.Fields
    lFlags = 0&
    avColName(j) = oField.Name
    avColType(j) = ColTypeToVB(oField.Type, lFlags)
    avColSize(j) = oField.DefinedSize
    If GetFieldProperty(oField, "ISAUTOINCREMENT") Then
      lFlags = lFlags Or ADOCOL_AUTOINCREMENT
    End If
    If oField.Attributes And adFldIsNullable Then
      lFlags = lFlags Or ADOCOL_ISNULLABLE
    End If
    avColFlags(j) = lFlags
    j = j + 1
  Next
  poRow.ArrayDefine avColName, avColType, avColSize, avColFlags

  DefineRowFromSet = True
  
DefineRowFromSet_Exit:
  Exit Function
  
DefineRowFromSet_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume DefineRowFromSet_Exit
End Function

#End If 'ROWLISTLIB

'
' 01.01.00 Transactions (customize according to your provider capacities)
'
Public Function ADORequireTransFeature(ByVal pfRequiredOrFail As Boolean) As Boolean
  ADORequireTransFeature = mfRequireTransFeature
End Function

Public Function ADOSetRequireTransFeature(ByVal pfRequiredOrFail As Boolean) As Boolean
  ADOSetRequireTransFeature = mfRequireTransFeature
  mfRequireTransFeature = pfRequiredOrFail
End Function

'If it'll be ever used, should be to clear transaction stack: Call ADOForcedTransactionStackCount(0)
Public Sub ADOForcedTransactionStackCount(ByVal plStackCount As Long)
  mlStackedTransCt = plStackCount
End Sub

Public Function ADOBeginStackedTrans(ByRef pcnConn As ADODB.Connection) As Boolean
  Dim vErr      As Variant
  On Error GoTo ADOBeginStackedTrans_Err
  ClearErr
  
  If mlStackedTransCt = 0& Then Call pcnConn.BeginTrans
  mlStackedTransCt = mlStackedTransCt + 1&
  
  ADOBeginStackedTrans = True
  Exit Function
ADOBeginStackedTrans_Err:
  If Err.Number = 3251 Then
    If Not mfRequireTransFeature Then
      Resume Next 'transanctions are not supported, let it go
    End If
  End If
  SetErr "ADOBeginStackedTrans", Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      On Error Resume Next
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
        msLastSQLState = vErr.SQLState
      Next
    End If
  End If
  Set pcnConn = Nothing
End Function

Public Function ADOIsTransPending() As Boolean
  ADOIsTransPending = CBool(mlStackedTransCt > 0&)
End Function

Public Function ADOEndStackedTrans(ByRef pcnConn As ADODB.Connection, ByVal fCommit As Boolean) As Boolean
  Dim vErr      As Variant
  On Error GoTo ADOEndStackedTrans_Err
  ClearErr
  
  mlStackedTransCt = mlStackedTransCt - 1&
  If mlStackedTransCt <= 0& Then  'think about the "<": it means that we'll let ADO fail in case of badly stacked trans
    If fCommit Then
      Call pcnConn.CommitTrans
    Else
      Call pcnConn.RollbackTrans
    End If
  End If
  'Correct counter if needed
  If mlStackedTransCt < 0& Then mlStackedTransCt = 0&
  
  ADOEndStackedTrans = True
  Exit Function
ADOEndStackedTrans_Err:
  If Err.Number = 3251 Then
    If Not mfRequireTransFeature Then
      Resume Next 'transanctions are not supported, let it go
    End If
  End If
  SetErr "ADOEndStackedTrans", Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    On Error Resume Next
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
        msLastSQLState = vErr.SQLState
      Next
    End If
  End If
  Set pcnConn = Nothing
End Function

Public Function GetFieldProperty(oField As ADODB.Field, ByVal psPropertyName As String) As Variant
  On Error Resume Next
  GetFieldProperty = oField.Properties(psPropertyName)
  If Err.Number Then
    GetFieldProperty = Null
  End If
End Function

Public Function ColTypeToVB(ByVal iADOType As Integer, ByRef lModifiableFlags As Long) As VbVarType
  'not exhaustive list
  Select Case iADOType
  Case adBinary, adVarBinary, adLongVarBinary
    ColTypeToVB = vbByte Or vbArray
    lModifiableFlags = lModifiableFlags Or ADOCOL_BLOB
  Case adTinyInt, adUnsignedTinyInt
    ColTypeToVB = vbByte
  Case adBoolean
    ColTypeToVB = vbBoolean
  Case adCurrency
    ColTypeToVB = vbCurrency
  Case adNumeric, adDecimal
    ColTypeToVB = vbDecimal
  Case adDouble
    ColTypeToVB = vbDouble
  Case adGUID
    ColTypeToVB = vbString
  Case adSmallInt, adUnsignedSmallInt
    ColTypeToVB = vbInteger
  Case adInteger, adUnsignedInt
    ColTypeToVB = vbLong
  Case adSingle
    ColTypeToVB = vbSingle
  Case adChar, adWChar, adVarChar, adVarWChar
    ColTypeToVB = vbString
  Case adLongVarWChar, adLongVarChar
    ColTypeToVB = vbString
    lModifiableFlags = lModifiableFlags Or ADOCOL_LONGTEXT
  Case adDate, adDBDate, adDBTimeStamp, adDBTime, adFileTime
    ColTypeToVB = vbDate
  Case Else
    ColTypeToVB = vbString
  End Select
End Function

Public Function ChainStmts(ByRef pcnConn As ADODB.Connection, ByRef psStmts As String, ByRef psStmtSep As String) As Boolean
  Dim fOK     As Boolean
  Dim iCount  As Long
  Dim i       As Long
  Dim asSQL() As String
  Dim fIgnoreErr  As Boolean
  
  iCount = SplitString(asSQL(), psStmts, psStmtSep)
  For i = 1& To iCount
    If Len(asSQL(i)) Then
      If (Left$(asSQL(i), 1) = "@") And (Left$(asSQL(i), 2) <> "@@") Then
        asSQL(i) = Right$(asSQL(i), Len(asSQL(i)) - 1)
        fIgnoreErr = True
      Else
        fIgnoreErr = False
      End If
      fOK = ADOExecSQL(pcnConn, asSQL(i))
      'Debug.Print "<"; asSQL(i); ">"
      If Not fOK Then
        If Not fIgnoreErr Then
          'Debug.Print "[FAILED] "; asSQL(i)
          'Debug.Print ADOLastErrDesc()
          'SetErr ADOLastErrCtx(), ADOLastErr(), ADOLastErrDesc() '& vbCrLf & asSQL(i)
          Exit Function
        End If
      End If
      'Debug.Print "...ok."
    End If
  Next i
  ChainStmts = True
End Function

Public Function ADOFieldExists(ByRef pcnConn As ADODB.Connection, ByVal psTableOrQueryName As String, ByVal psFieldname As String) As Boolean
  Const LOCAL_ERR_CTX As String = "ADOFieldExists"
  On Error GoTo ADOFieldExists_Err
  ClearErr
  
  Dim SQL         As String
  Dim rsData      As ADODB.Recordset
  Dim oFld        As ADODB.Field
  Dim vErr        As Variant
  
  SQL = "SELECT * FROM [" & psTableOrQueryName & "] WHERE 1=2"  'We just want the column names
  Set rsData = New ADODB.Recordset
  rsData.Open SQL, pcnConn, adOpenStatic, adLockReadOnly
  On Error Resume Next
  Set oFld = rsData.Fields(psFieldname)
  If Err.Number = 0 Then
    ADOFieldExists = True
  End If
  Set oFld = Nothing
  On Error GoTo ADOFieldExists_Err
  rsData.Close
  
ADOFieldExists_Exit:
  Set rsData = Nothing
  Exit Function

ADOFieldExists_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Resume ADOFieldExists_Exit
End Function

Public Function ADOIndexExists(ByRef pcnConn As ADODB.Connection, ByVal psTableName As String, ByVal psIndexName As String) As Boolean
  Const LOCAL_ERR_CTX As String = "ADOIndexExists"
  On Error GoTo ADOIndexExists_Err
  ClearErr
  
  Dim vErr  As Variant
  Dim cat   As ADOX.Catalog
  Dim tbl   As ADOX.Table
  Dim idx   As ADOX.Index
  
  Set cat = New ADOX.Catalog
  cat.ActiveConnection = pcnConn
  
  Set tbl = cat.Tables(psTableName)
  
  For Each idx In tbl.Indexes
    If StrComp(idx.Name, psIndexName, vbTextCompare) = 0 Then
      ADOIndexExists = True
      GoTo ADOIndexExists_Exit
    End If
  Next idx
  
ADOIndexExists_Exit:
  Set idx = Nothing
  Set tbl = Nothing
  Set cat = Nothing
  Exit Function

ADOIndexExists_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Resume ADOIndexExists_Exit
End Function

Public Function ADOGetFieldType(ByRef pcnConn As ADODB.Connection, ByVal psTableOrQueryName As String, ByVal psFieldname As String) As ADODB.DataTypeEnum
  Const LOCAL_ERR_CTX As String = "ADOGetFieldType"
  On Error GoTo ADOGetFieldType_Err
  ClearErr
  
  Dim SQL         As String
  Dim rsData      As ADODB.Recordset
  Dim oFld        As ADODB.Field
  Dim vErr        As Variant
  
  ADOGetFieldType = ADODB.DataTypeEnum.adVarChar  'by default
  
  SQL = "SELECT * FROM [" & psTableOrQueryName & "] WHERE 1=2"  'We just want the column names
  Set rsData = New ADODB.Recordset
  rsData.Open SQL, pcnConn, adOpenStatic, adLockReadOnly
  On Error Resume Next
  Set oFld = rsData.Fields(psFieldname)
  If Err.Number = 0 Then
    ADOGetFieldType = oFld.Type
  End If
  Set oFld = Nothing
  On Error GoTo ADOGetFieldType_Err
  rsData.Close
  
ADOGetFieldType_Exit:
  Set rsData = Nothing
  Exit Function

ADOGetFieldType_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  Resume ADOGetFieldType_Exit
End Function

Public Function ADOGetFieldProperty( _
  ByRef pcnConn As ADODB.Connection, _
  ByVal psTableName As String, _
  ByVal psFieldname As String, _
  ByVal psPropertyName As String) As Variant
  Const LOCAL_ERR_CTX As String = "ADOGetFieldProperty"
  On Error GoTo ADOGetFieldProperty_Err
  ClearErr
  
  Dim vErr        As Variant
  Dim fOK         As Boolean
  Dim oCat        As New ADOX.Catalog
  Dim oTbl        As ADOX.Table
  Dim vRet        As Variant
  
  fOK = True
  vRet = Null
  
  Set oCat.ActiveConnection = pcnConn
  Set oTbl = oCat.Tables(psTableName)
  
'  Dim i As Integer
'  For i = 0 To oTbl.Columns(psFieldname).Properties.Count - 1
'    Debug.Print oTbl.Columns(psFieldname).Properties(i).Name
'  Next i
  
  vRet = oTbl.Columns(psFieldname).Properties(psPropertyName)
  
ADOGetFieldProperty_Exit:
  Set oTbl = Nothing
  Set oCat.ActiveConnection = Nothing
  ADOGetFieldProperty = vRet
  Exit Function

ADOGetFieldProperty_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  fOK = False
  Resume ADOGetFieldProperty_Exit
End Function

Public Function ADOSetFieldProperty( _
  ByRef pcnConn As ADODB.Connection, _
  ByVal psTableName As String, _
  ByVal psFieldname As String, _
  ByVal psPropertyName As String, _
  ByVal pvValue As Variant) As Boolean
  Const LOCAL_ERR_CTX As String = "ADOSetFieldProperty"
  On Error GoTo ADOSetFieldProperty_Err
  ClearErr
  
  Dim vErr        As Variant
  Dim fOK         As Boolean
  Dim oCat        As New ADOX.Catalog
  Dim oTbl        As ADOX.Table
  
  fOK = True
  
  Set oCat.ActiveConnection = pcnConn
  Set oTbl = oCat.Tables(psTableName)
  
  oTbl.Columns(psFieldname).Properties(psPropertyName) = pvValue
  
ADOSetFieldProperty_Exit:
  Set oTbl = Nothing
  Set oCat.ActiveConnection = Nothing
  ADOSetFieldProperty = fOK
  Exit Function

ADOSetFieldProperty_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  fOK = False
  Resume ADOSetFieldProperty_Exit
End Function

Public Function ADOGetFieldAttribs( _
  ByRef pcnConn As ADODB.Connection, _
  ByVal psTableName As String, _
  ByVal psFieldname As String) As Long
  Const LOCAL_ERR_CTX As String = "ADOGetFieldAttribs"
  On Error GoTo ADOGetFieldAttribs_Err
  ClearErr
  
  Dim vErr        As Variant
  Dim fOK         As Boolean
  Dim oCat        As New ADOX.Catalog
  Dim oTbl        As ADOX.Table
  Dim lRet        As Long
  
  fOK = True
  
  Set oCat.ActiveConnection = pcnConn
  Set oTbl = oCat.Tables(psTableName)
  
  lRet = oTbl.Columns(psFieldname).Attributes
  
ADOGetFieldAttribs_Exit:
  Set oTbl = Nothing
  Set oCat.ActiveConnection = Nothing
  ADOGetFieldAttribs = lRet
  Exit Function

ADOGetFieldAttribs_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  fOK = False
  Resume ADOGetFieldAttribs_Exit
End Function

Public Function ADOSetFieldNullable( _
  ByRef pcnConn As ADODB.Connection, _
  ByVal psTableName As String, _
  ByVal psFieldname As String, _
  ByVal pfNullable As Boolean) As Boolean
  Const LOCAL_ERR_CTX As String = "ADOSetFieldNullable"
  On Error GoTo ADOSetFieldNullable_Err
  ClearErr
  
  Dim vErr        As Variant
  Dim fOK         As Boolean
  Dim oCat        As New ADOX.Catalog
  Dim oTbl        As ADOX.Table
  
  fOK = True
  
  Set oCat.ActiveConnection = pcnConn
  Set oTbl = oCat.Tables(psTableName)
  
  If pfNullable Then
    oTbl.Columns(psFieldname).Attributes = oTbl.Columns(psFieldname).Attributes Or adColNullable
  Else
    oTbl.Columns(psFieldname).Attributes = oTbl.Columns(psFieldname).Attributes And (Not adColNullable)
  End If
  
ADOSetFieldNullable_Exit:
  Set oTbl = Nothing
  Set oCat.ActiveConnection = Nothing
  ADOSetFieldNullable = fOK
  Exit Function

ADOSetFieldNullable_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  If Not pcnConn Is Nothing Then
    If pcnConn.Errors.Count Then
      msErr = ""
      mlErr = Err.Number Xor vbObjectError
      For Each vErr In pcnConn.Errors
        If Len(msErr) Then msErr = msErr & vbCrLf
        msErr = msErr & (vErr.Number Xor vbObjectError) & ": " & vErr.Description
      Next
    End If
  End If
  fOK = False
  Resume ADOSetFieldNullable_Exit
End Function

#If TWINBASIC Then
End Module
#End If
