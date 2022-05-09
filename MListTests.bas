Attribute VB_Name = "MListTests"
#If TWINBASIC Then
[ TestFixture ]
Private Module MListTests
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Private mlstSuppliers   As CList
Private mlstShippers    As CList
Private mlstProducts    As CList

'
' Module dedicated to the CList tests
'

Public Sub Test_RunAllTests()
  On Error GoTo Test_RunAllTests_Err
  
  Test_BeginSuite "MListTests"

'  TestListAddRemoveData
'  TestListSortMethods
  TestListSortWithBlanks
'  TestListFindMethods
'  TestProductListToJson
'  TestListGroupBy1
'  TestListGroupByWithOneElement
'  TestListGroupByProducts
'  TestJsonTestFiles
  
Test_RunAllTests_Exit:
  Test_EndSuite
  Exit Sub

Test_RunAllTests_Err:
  Debug.Print "Test_RunAllTests() failed: " & Err.Description
  Resume Test_RunAllTests_Exit
End Sub

Private Function SameRowValues(ByRef poRow1 As CRow, ByVal poRow2 As CRow) As Boolean
  Dim iCount1   As Integer
  Dim iCount2   As Integer
  Dim iCol      As Integer
  
  iCount1 = poRow1.ColCount
  iCount2 = poRow2.ColCount
  
  For iCol = 1 To iCount1
    If poRow1(iCol) <> poRow2(iCol) Then
      SameRowValues = False
      Exit Function
    End If
  Next iCol
  
  SameRowValues = True
End Function

'Adding, updating and removing data to a list
#If TWINBASIC Then
[ TestCase ]
#End If
Sub TestListAddRemoveData()
  Const LOCAL_ERR_CTX As String = "TestListAddRemoveData"
  On Error GoTo TestListAddRemoveData_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Adding, updating and removing data to a list"
  
  'Define and populate lists we're going to use.
  'We'll use the Supplier list.
  If Not CreateSuppliersList() Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to create test suppliers list"
    GoTo TestListAddRemoveData_Exit
  End If
  'And we'll use the Shipper's list
  If Not CreateShippersList() Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to create test shippers list"
    GoTo TestListAddRemoveData_Exit
  End If
  'We can define a row from the list:
  Dim oRow      As New CRow
  mlstSuppliers.DefineRow oRow
  'Populate the row
  oRow("SupplierID") = 30&
  oRow("CompanyName") = "DIS"
  oRow("ContactName") = "Francesco Foti"
  oRow("ContactTitle") = "Mr."
  oRow("Address") = "47, avenue Blanc"
  oRow("City") = "Geneva"
  oRow("Region") = "Geneva"
  oRow("PostalCode") = "1202"
  oRow("Country") = "Switzerland"
  oRow("Phone") = "+41 (22) 555 55 48"
  oRow("Fax") = "+41 (22) 555 55 55"
  oRow("HomePage") = "devinfo.net#http://www.devinfo.net#"
  'Add the row to the list, in first position and dump the list to check
  mlstSuppliers.AddRow oRow, plInsertBefore:=1&
  'ListDump mlstSuppliers, "Suppliers"
  Test_Value "DIS", mlstSuppliers("CompanyName", 1&), "Invalid list position for new supplier"
  
  'We then take a row from the Shipper's list, which has a different
  'definition, and we add it to the supplier list, just to illustrate
  'the column matching algorithm. Columns that are not matched will be Null.
  Dim oShippersRow    As CRow
  Set oShippersRow = mlstShippers.row(2)  'This creates a CRow object and *copies* values into it.
  'Add the row to the list, in 2nd position and dump the list to check
  mlstSuppliers.AddRow oShippersRow, plInsertBefore:=2&
  ListDump mlstSuppliers.row(2) 'This creates a temporary reference (on a copy) that lives during the call
  'ListDump mlstSuppliers, "Suppliers"
  Set oShippersRow = Nothing
  'The SupplierID is missing for row 2, so we assign a value to the cell
  Test_Value Null, mlstSuppliers("SupplierID", 2&), "New supplier ID should be null"
  mlstSuppliers("SupplierID", 2&) = 31
  'And we just dump the modified row
  ListDump mlstSuppliers.row(2), "New, copied, supplier row #2" 'This creates a temporary reference that lives during the call
  
  'To test and demonstrate the assign row method, we'll copy a row on another,
  'creating a duplicate-
  mlstSuppliers.GetRow oRow, 1
  mlstSuppliers.AssignRow 2, oRow
  'row 2 now equals row 1
  Test_Value True, SameRowValues(mlstSuppliers.row(1&), mlstSuppliers.row(2&)), "Row values should be exactly the same"
  
  'Now we just changed the first two columns of row 0.
  'Note that using the AssignValues method, implies that
  'we know and respect the list columns order.
  mlstSuppliers.AssignValues 1, 32&, "Any company"
  'And we just dump the modified row
  ListDump mlstSuppliers.row(1)
  Test_Value 32&, mlstSuppliers("SupplierID", 1), "Updated ID mismatch"
  
TestListAddRemoveData_Exit:
  'Destroy lists
  Set mlstShippers = Nothing
  Set mlstSuppliers = Nothing
  
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestListAddRemoveData_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListAddRemoveData_Exit
  Resume
End Sub

'Test Sort methods
#If TWINBASIC Then
[ TestCase ]
#End If
Sub TestListSortMethods()
  Const LOCAL_ERR_CTX As String = "TestListSortMethods"
  On Error GoTo TestListSortMethods_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test Sort methods"
  
  Dim i     As Long
  
  'Create a products list
  If Not CreateProductsList() Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Couldn't create product list"
    GoTo TestListSortMethods_Exit
  End If
  'Sort on a column, ascending
  mlstProducts.Sort "CompanyName+"
  'ListDump mlstProducts, "Products, sorted on 'CompanyName+'"
  'Check
  For i = mlstProducts.Count To 2& Step -1&
    If StrComp(mlstProducts("CompanyName", i), mlstProducts("CompanyName", i - 1&), vbTextCompare) < 0 Then
      Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to sorting on [CompanyName+]"
      GoTo TestListSortMethods_Exit
    End If
  Next i
  'Sort descending, use "-" sort indicator
  mlstProducts.Sort "CompanyName-"
  'Check
  For i = mlstProducts.Count To 2& Step -1&
    If StrComp(mlstProducts("CompanyName", i), mlstProducts("CompanyName", i - 1&), vbTextCompare) > 0 Then
      Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to sorting on [CompanyName-]"
      GoTo TestListSortMethods_Exit
    End If
  Next i

  'Sort descending, case sensitive, banging the column name
  mlstProducts.Sort "!CompanyName-"
  'Check
  For i = mlstProducts.Count To 2& Step -1&
    If StrComp(mlstProducts("CompanyName", i), mlstProducts("CompanyName", i - 1&), vbBinaryCompare) > 0 Then
      Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to sorting on [!CompanyName-]"
      GoTo TestListSortMethods_Exit
    End If
  Next i

  'Sort on multiple columns
  mlstProducts.Sort "!CompanyName-,ProductName"
  'Check
  'ListDump mlstProducts, "Products, sorted on '[!CompanyName-,ProductName]'", "CompanyName:40;ProductName:40"
  'When sorted, first product id is 48, last is 38 (verified visually on a correct sort, uncomment previous list dump)
  If (mlstProducts("ProductID", 1&) <> 48&) Or (mlstProducts("ProductID", mlstProducts.Count) <> 38&) Then
    Call OpenTraceOutputFile
    ListDump mlstProducts, "Products, sorted on '[!CompanyName-,ProductName]'", "CompanyName:40;ProductName:40"
    CloseTraceOutputFile
    ViewTraceOutputFile
    Test_SetSuccess LOCAL_ERR_CTX, False, "Sorting on [!CompanyName-,ProductName] FAILED"
    GoTo TestListSortMethods_Exit
  End If
  
  i = mlstProducts.Find("ProductID", 55&)
  If i > 0 Then
    mlstProducts("CompanyName", i) = "ma maison"
    mlstProducts.Sort "!CompanyName-,ProductName"
    If mlstProducts("ProductID", 1&) <> 55& Then
      Call OpenTraceOutputFile
      ListDump mlstProducts, "Products, sorted on '[!CompanyName-,ProductName]'", "CompanyName:40;ProductName:40"
      CloseTraceOutputFile
      ViewTraceOutputFile
      Test_SetSuccess LOCAL_ERR_CTX, False, "lowercasing a company name should make it first element"
    End If
  End If

  Test_SetSuccess LOCAL_ERR_CTX, True
  
TestListSortMethods_Exit:
  Set mlstProducts = Nothing
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestListSortMethods_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListSortMethods_Exit
End Sub

Sub TestListSortWithBlanks()
  Const LOCAL_ERR_CTX As String = "TestListSortWithBlanks"
  On Error GoTo TestListSortWithBlanks_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test Sort methods"
  
  Dim i     As Long
  Dim lstData     As CList
  
  'A list with a empty strings
  Set lstData = New CList
  lstData.ArrayDefine Array("text", "double"), Array(vbString, vbDouble)
  lstData.AddValues "klmnop", 1.5
  lstData.AddValues "zabcde", -25
  lstData.AddValues "hijklm", 0
  lstData.AddValues "nomopq", 934567.456787
  lstData.AddValues "", 5.5
  lstData.AddValues "", 12
  lstData.AddValues "", 657
  lstData.AddValues "", 43
  lstData.AddValues "abcdef", 3.5
  lstData.Sort "text,double"
  ListDump lstData, "Correct sort on 'text,double'", "text:25"
  
  For i = 1 To lstData.Count - 1
    If (lstData("text", i) > lstData("text", i + 1)) Or _
       ((lstData("text", i) = lstData("text", i + 1)) And (lstData("double", i) > lstData("double", i + 1))) Then
      Test_Comment "Entry #" & i & " is not in correct sort order"
      Test_SetSuccess LOCAL_ERR_CTX, False
      GoTo TestListSortWithBlanks_Exit
    End If
  Next i
  
  'A list with null values
  Set lstData = New CList
  lstData.ArrayDefine Array("text", "double"), Array(vbString, vbDouble)
  lstData.AddValues "klmnop", 1.5
  lstData.AddValues "zabcde", Null
  lstData.AddValues "hijklm", 0
  lstData.AddValues "nomopq", 934567.456787
  lstData.AddValues "", 5.5
  lstData.AddValues Null, 12
  lstData.AddValues "", 657
  lstData.AddValues "", 43
  lstData.AddValues "", Null
  lstData.AddValues Null, Null
  lstData.AddValues "abcdef", 3.5
  lstData.Sort "text,double"
  
  For i = 1 To lstData.Count - 1
    If lstData("text", i) > lstData("text", i + 1) Then
      Test_Comment "Entry #" & i & " is not in correct sort order"
      ListDump lstData, "incorrect sort on 'text,double'", "text:25"
      Test_SetSuccess LOCAL_ERR_CTX, False
      GoTo TestListSortWithBlanks_Exit
    End If
  Next i
  ListDump lstData, "Correct sort on 'text,double' with nulls", "text:25"
  
  'A resembling the problematic one in ffchilkatldr_test working project
  Set lstData = New CList
  lstData.ArrayDefine Array("value", "block"), Array(vbString, vbDouble)
  lstData.AddValues "~$ Feedtrade Share sale and purchase agreement GY and YL 04.03.2014", 8709
  lstData.AddValues "", 9969
  lstData.AddValues ChrW$(1050) & ChrW$(1086) & ChrW$(1087) & ChrW$(1103) & ChrW$(1050) & " MAY BACK UP (EASY WAY) (1)", 3224
  lstData.AddValues "nomopq", 934567.456787
  lstData.AddValues "", 5.5
  lstData.AddValues Null, 12
  lstData.AddValues "", 657
  lstData.AddValues "", 43
  lstData.AddValues "", Null
  lstData.AddValues Null, Null
  lstData.AddValues "abcdef", 3.5
  lstData.Sort "value,block"
  
  For i = 1 To lstData.Count - 1
    If lstData("text", i) > lstData("text", i + 1) Then
      Test_Comment "Entry #" & i & " is not in correct sort order"
      ListDump lstData, "incorrect sort on 'text,double'", "text:25"
      Test_SetSuccess LOCAL_ERR_CTX, False
      GoTo TestListSortWithBlanks_Exit
    End If
  Next i
  ListDump lstData, "Correct sort on 'text,double' with nulls", "text:25"
   
  Test_SetSuccess LOCAL_ERR_CTX, True
  
TestListSortWithBlanks_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestListSortWithBlanks_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListSortWithBlanks_Exit
End Sub

'Test find methods
#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestListFindMethods()
  Const LOCAL_ERR_CTX As String = "TestListFindMethods"
  On Error GoTo TestListFindMethods_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test Find methods"
  
  Dim sFindWhat     As String
  Dim lRow          As Long
  
  'Create a products list
  If Not CreateProductsList() Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Couldn't create product list"
    GoTo TestListFindMethods_Exit
  End If
  
  'Sort case insensitive ascending sort order, on ProductName field
  mlstProducts.Sort "ProductName"
  'ListDump mlstProducts, "Products, sorted on 'ProductName'"
  'We have sorted, using *case insensitive* search, so we can search without
  'worrying for the letter case of our search criteria:
  'Let's find "Gnocchi di nonna Alice"...
  sFindWhat = "Gnocchi di nonna Alice"
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] found in column 'ProductName' with list sorted on it case insensitive"
  
  sFindWhat = UCase$("Gnocchi di nonna Alice")
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] found in column 'ProductName' with list sorted on it case insensitive"
  
  'But if we sort, specifying a case sensitive sort, then we have
  'to give the exact value to find back our data:
  mlstProducts.Sort "!ProductName"
  sFindWhat = "!" & UCase$("Gnocchi di nonna alice")
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_Value 0, lRow, "[" & sFindWhat & "] found in column 'ProductName' with list sorted on it case sensitive"
  
  'Search specifying the root of the search term(s)
  mlstProducts.Sort "ProductName"
  sFindWhat = "Gnocchi*"
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] found in column 'ProductName' with list sorted on it case insensitive"
  
  'Search specifying a suffix
  sFindWhat = "*nonna Alice"
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] found in column 'ProductName' with list sorted on it case insensitive"

  'Search specifying a suffix, but with incorrect case (sorted case sensitive)
  mlstProducts.Sort "!ProductName"
  sFindWhat = "!*NONNA Alice" 'Search expression case matters (begins with a bang '!')
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_Value 0, lRow, "[" & sFindWhat & "] should not have been found in column 'ProductName' with list sorted case sensitive"
  'Do it again, but before sort indicating that case doesn't matter.
  mlstProducts.Sort "ProductName"
  sFindWhat = "*NONNA Alice"  'Search expr case doesn't matter && sorted case insensitive
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] found in column 'ProductName' with list sorted on it case insensitive"
  'now it will work
  sFindWhat = "!*NONNA Alice"
  lRow = mlstProducts.Find("ProductName", sFindWhat)
  Test_Value 0, lRow, "[" & sFindWhat & "] should not have been found in column 'ProductName' with list sorted case insensitive"
  
  'Now let's try to find ProductName "Chocolade", but searching only
  'in rows where the CategoryName is "Confections".
  'There are many ways to do that, but one of the fastest is to sort
  'on the category name, find the first line for which the category
  'is "Confections" and then sequentially search for "Chocolade".
  'To benefit from the list object facilities, we first use FindFirst,
  'and then we use a simple find.
  mlstProducts.Sort "CategoryName"
  sFindWhat = "Confections"
  lRow = mlstProducts.FindFirst("CategoryName", sFindWhat)
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] found in 'CategoryName'"
  'This will be a sequential search, for 2 reasons:
  ' 1. A joker is used
  ' 2. We're not searching on a sorted column
  'Note: we'll use a column number notation, instead of the column name,
  'just to give it a try. ProductName is in column 2.
  sFindWhat = "Chocola?e"
  lRow = mlstProducts.Find("#2", sFindWhat, lRow)  'We use a joker
  'With a named column: lRow = mlstProducts.Find("ProductName", "Chocola?e", lRow)  'We use a joker
  Test_ValueNotEqual 0, lRow, "[" & sFindWhat & "] should found in column 'CategoryName' with list sorted case insensitive"
  If lRow Then
    'we have to test again the category, as we may have gone too far.
    'We would be faster if we knew the last row index which category is "Confections".
    Test_Value "Confections", mlstProducts("CategoryName", lRow), "'CategoryName' is 'Confections'"
  End If
  
  'Remove duplicates test.
  'Note that this could be useful in this example, to find
  'the distinct number and names of categories.
  mlstProducts.RemoveDuplicates
  ListDump mlstProducts, "Products w/o duplicates on category"
  
  'Define another list as the mlstProducts list
  Dim oNewList As New CList
  mlstProducts.DefineList oNewList, Null, ""
  ListDump oNewList
  'That was fine, but now, copy the entire list
  oNewList.CopyFrom mlstProducts
  ListDump oNewList, "Copy"
  
TestListFindMethods_Exit:
  Set mlstProducts = Nothing
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestListFindMethods_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListFindMethods_Exit
End Sub

'
' Creating the test lists
'

Function CreateSuppliersList() As Boolean
  'Setup error trap
  On Error GoTo CreateSuppliersList_Err
  'create a new list object
  Set mlstSuppliers = New CList
  mlstSuppliers.ArrayDefine Array("SupplierID", "CompanyName", "ContactName", "ContactTitle", _
                                  "Address", "City", "Region", "PostalCode", _
                                  "Country", "Phone", "Fax", "HomePage"), _
                           Array(vbLong, vbString, vbString, vbString, _
                                 vbString, vbString, vbString, vbString, _
                                 vbString, vbString, vbString, vbString)
  'Add values to our list.
  'The AddValues lines have been generated from the Supplier table of the NWIND.mdb
  With mlstSuppliers
    .AddValues 1, "Exotic Liquids", "Charlotte Cooper", "Purchasing Manager", "49 Gilbert St.", "London", Null, "EC1 4SD", "UK", "(171) 555-2222", Null, Null
    .AddValues 2, "New Orleans Cajun Delights", "Shelley Burke", "Order Administrator", "P.O. Box 78934", "New Orleans", "LA", "70117", "USA", "(100) 555-4822", Null, "#CAJUN.HTM#"
    .AddValues 3, "Grandma Kelly's Homestead", "Regina Murphy", "Sales Representative", "707 Oxford Rd.", "Ann Arbor", "MI", "48104", "USA", "(313) 555-5735", "(313) 555-3349", Null
    .AddValues 4, "Tokyo Traders", "Yoshi Nagase", "Marketing Manager", "9-8 Sekimai" & vbCrLf & "Musashino-shi", "Tokyo", Null, "100", "Japan", "(03) 3555-5011", Null, Null
    .AddValues 5, "Cooperativa de Quesos 'Las Cabras'", "Antonio del Valle Saavedra ", "Export Administrator", "Calle del Rosal 4", "Oviedo", "Asturias", "33007", "Spain", "(98) 598 76 54", Null, Null
    .AddValues 6, "Mayumi's", "Mayumi Ohno", "Marketing Representative", "92 Setsuko" & vbCrLf & "Chuo-ku", "Osaka", Null, "545", "Japan", "(06) 431-7877", Null, "Mayumi's (on the World Wide Web)#http://www.microsoft.com/accessdev/sampleapps/mayumi.htm#"
    .AddValues 7, "Pavlova, Ltd.", "Ian Devling", "Marketing Manager", "74 Rose St." & vbCrLf & "Moonie Ponds", "Melbourne", "Victoria", "3058", "Australia", "(03) 444-2343", "(03) 444-6588", Null
    .AddValues 8, "Specialty Biscuits, Ltd.", "Peter Wilson", "Sales Representative", "29 King's Way", "Manchester", Null, "M14 GSD", "UK", "(161) 555-4448", Null, Null
    .AddValues 9, "PB Knäckebröd AB", "Lars Peterson", "Sales Agent", "Kaloadagatan 13", "Göteborg", Null, "S-345 67", "Sweden ", "031-987 65 43", "031-987 65 91", Null
    .AddValues 10, "Refrescos Americanas LTDA", "Carlos Diaz", "Marketing Manager", "Av. das Americanas 12.890", "São Paulo", Null, "5442", "Brazil", "(11) 555 4640", Null, Null
    .AddValues 11, "Heli Süßwaren GmbH & Co. KG", "Petra Winkler", "Sales Manager", "Tiergartenstraße 5", "Berlin", Null, "10785", "Germany", "(010) 9984510", Null, Null
    .AddValues 12, "Plutzer Lebensmittelgroßmärkte AG", "Martin Bein", "International Marketing Mgr.", "Bogenallee 51", "Frankfurt", Null, "60439", "Germany", "(069) 992755", Null, "Plutzer (on the World Wide Web)#http://www.microsoft.com/accessdev/sampleapps/plutzer.htm#"
    .AddValues 13, "Nord-Ost-Fisch Handelsgesellschaft mbH", "Sven Petersen", "Coordinator Foreign Markets", "Frahmredder 112a", "Cuxhaven", Null, "27478", "Germany", "(04721) 8713", "(04721) 8714", Null
    .AddValues 14, "Formaggi Fortini s.r.l.", "Elio Rossi", "Sales Representative", "Viale Dante, 75", "Ravenna", Null, "48100", "Italy", "(0544) 60323", "(0544) 60603", "#FORMAGGI.HTM#"
    .AddValues 15, "Norske Meierier", "Beate Vileid", "Marketing Manager", "Hatlevegen 5", "Sandvika", Null, "1320", "Norway", "(0)2-953010", Null, Null
    .AddValues 16, "Bigfoot Breweries", "Cheryl Saylor", "Regional Account Rep.", "3400 - 8th Avenue" & vbCrLf & "Suite 210", "Bend", "OR", "97101", "USA", "(503) 555-9931", Null, Null
    .AddValues 17, "Svensk Sjöföda AB", "Michael Björn", "Sales Representative", "Brovallavägen 231", "Stockholm", Null, "S-123 45", "Sweden", "08-123 45 67", Null, Null
    .AddValues 18, "Aux joyeux ecclésiastiques", "Guylène Nodier", "Sales Manager", "203, Rue des Francs-Bourgeois", "Paris", Null, "75004", "France", "(1) 03.83.00.68", "(1) 03.83.00.62", Null
    .AddValues 19, "New England Seafood Cannery", "Robb Merchant", "Wholesale Account Agent", "Order Processing Dept." & vbCrLf & "2100 Paul Revere Blvd.", "Boston", "MA", "02134", "USA", "(617) 555-3267", "(617) 555-3389", Null
    .AddValues 20, "Leka Trading", "Chandra Leka", "Owner", "471 Serangoon Loop, Suite #402", "Singapore", Null, "0512", "Singapore", "555-8787", Null, Null
    .AddValues 21, "Lyngbysild", "Niels Petersen", "Sales Manager", "Lyngbysild" & vbCrLf & "Fiskebakken 10", "Lyngby", Null, "2800", "Denmark", "43844108", "43844115", Null
    .AddValues 22, "Zaanse Snoepfabriek", "Dirk Luchte", "Accounting Manager", "Verkoop" & vbCrLf & "Rijnweg 22", "Zaandam", Null, "9999 ZZ", "Netherlands", "(12345) 1212", "(12345) 1210", Null
    .AddValues 23, "Karkki Oy", "Anne Heikkonen", "Product Manager", "Valtakatu 12", "Lappeenranta", Null, "53120", "Finland", "(953) 10956", Null, Null
    .AddValues 24, "G'day, Mate", "Wendy Mackenzie", "Sales Representative", "170 Prince Edward Parade" & vbCrLf & "Hunter's Hill", "Sydney", "NSW", "2042", "Australia", "(02) 555-5914", "(02) 555-4873", "G'day Mate (on the World Wide Web)#http://www.microsoft.com/accessdev/sampleapps/gdaymate.htm#"
    .AddValues 25, "Ma Maison", "Jean-Guy Lauzon", "Marketing Manager", "2960 Rue St. Laurent", "Montréal", "Québec", "H1J 1C3", "Canada", "(514) 555-9022", Null, Null
    .AddValues 26, "Pasta Buttini s.r.l.", "Giovanni Giudici", "Order Administrator", "Via dei Gelsomini, 153", "Salerno", Null, "84100", "Italy", "(089) 6547665", "(089) 6547667", Null
    .AddValues 27, "Escargots Nouveaux", "Marie Delamare", "Sales Manager", "22, rue H. Voiron", "Montceau", Null, "71300", "France", "85.57.00.07", Null, Null
    .AddValues 28, "Gai pâturage", "Eliane Noz", "Sales Representative", "Bat. B" & vbCrLf & "3, rue des Alpes", "Annecy", Null, "74000", "France", "38.76.98.06", "38.76.98.58", Null
    .AddValues 29, "Forêts d'érables", "Chantal Goulet", "Accounting Manager", "148 rue Chasseur", "Ste-Hyacinthe", "Québec", "J2S 7S8", "Canada", "(514) 555-2955", "(514) 555-2921", Null
  End With
  
  CreateSuppliersList = True
  Exit Function
CreateSuppliersList_Err:
  OutputLn "Error creating suppliers list: " & Err.Description
  Set mlstSuppliers = Nothing
End Function

Function CreateShippersList() As Boolean
  'Setup error trap
  On Error GoTo CreateShippersList_Err
  'create a new list object
  Set mlstShippers = New CList
  mlstShippers.ArrayDefine Array("ShipperID", "CompanyName", "Phone"), _
                           Array(vbLong, vbString, vbString)
  'Add values to our list.
  'The AddValues lines have been generated from the Shipper table of the NWIND.mdb
  With mlstShippers
    .AddValues 1, "Speedy Express", "(503) 555-9831"
    .AddValues 2, "United Package", "(503) 555-3199"
    .AddValues 3, "Federal Shipping", "(503) 555-9931"
  End With
  
  CreateShippersList = True
  Exit Function
CreateShippersList_Err:
  OutputLn "Error creating Shippers list: " & Err.Description
  Set mlstShippers = Nothing
End Function

Function CreateProductsList() As Boolean
  'Setup error trap
  On Error GoTo CreateProductsList_Err
  'create a new list object
  Set mlstProducts = New CList
  mlstProducts.ArrayDefine Array("ProductID", "ProductName", "SupplierID", "CompanyName", _
                                 "CategoryID", "CategoryName", "QuantityPerUnit", "UnitPrice", _
                                 "UnitsInStock", "UnitsOnOrder", "ReorderLevel", _
                                 "Discontinued"), _
                           Array(vbLong, vbString, vbLong, vbString, _
                                 vbLong, vbString, vbString, vbCurrency, _
                                 vbInteger, vbInteger, vbInteger, _
                                 vbBoolean)
  'Add values to our list.
  'The AddValues lines have been generated by a custom query on the NWIND.mdb
  With mlstProducts
    .AddValues 1, "Chai", 1, "Exotic Liquids", 1, "Beverages", "10 boxes x 20 bags", 18, 39, 0, 10, False
    .AddValues 2, "Chang", 1, "Exotic Liquids", 1, "Beverages", "24 - 12 oz bottles", 19, 17, 40, 25, False
    .AddValues 24, "Guaraná Fantástica", 10, "Refrescos Americanas LTDA", 1, "Beverages", "12 - 355 ml cans", 4.5, 20, 0, 0, True
    .AddValues 34, "Sasquatch Ale", 16, "Bigfoot Breweries", 1, "Beverages", "24 - 12 oz bottles", 14, 111, 0, 15, False
    .AddValues 35, "Steeleye Stout", 16, "Bigfoot Breweries", 1, "Beverages", "24 - 12 oz bottles", 18, 20, 0, 15, False
    .AddValues 38, "C�te de Blaye", 18, "Aux joyeux eccl�siastiques", 1, "Beverages", "12 - 75 cl bottles", 263.5, 17, 0, 15, False
    .AddValues 39, "Chartreuse verte", 18, "Aux joyeux eccl�siastiques", 1, "Beverages", "750 cc per bottle", 18, 69, 0, 5, False
    .AddValues 43, "Ipoh Coffee", 20, "Leka Trading", 1, "Beverages", "16 - 500 g tins", 46, 17, 10, 25, False
    .AddValues 67, "Laughing Lumberjack Lager", 16, "Bigfoot Breweries", 1, "Beverages", "24 - 12 oz bottles", 14, 52, 0, 10, False
    .AddValues 70, "Outback Lager", 7, "Pavlova, Ltd.", 1, "Beverages", "24 - 355 ml bottles", 15, 15, 10, 30, False
    .AddValues 75, "Rhönbräu Klosterbier", 12, "Plutzer Lebensmittelgroßmärkte AG", 1, "Beverages", "24 - 0.5 l bottles", 7.75, 125, 0, 25, False
    .AddValues 76, "Lakkalikööri", 23, "Karkki Oy", 1, "Beverages", "500 ml", 18, 57, 0, 20, False
    .AddValues 3, "Aniseed Syrup", 1, "Exotic Liquids", 2, "Condiments", "12 - 550 ml bottles", 10, 13, 70, 25, False
    .AddValues 4, "Chef Anton's Cajun Seasoning", 2, "New Orleans Cajun Delights", 2, "Condiments", "48 - 6 oz jars", 22, 53, 0, 0, False
    .AddValues 5, "Chef Anton's Gumbo Mix", 2, "New Orleans Cajun Delights", 2, "Condiments", "36 boxes", 21.35, 0, 0, 0, True
    .AddValues 6, "Grandma's Boysenberry Spread", 3, "Grandma Kelly's Homestead", 2, "Condiments", "12 - 8 oz jars", 25, 120, 0, 25, False
    .AddValues 8, "Northwoods Cranberry Sauce", 3, "Grandma Kelly's Homestead", 2, "Condiments", "12 - 12 oz jars", 40, 6, 0, 0, False
    .AddValues 15, "Genen Shouyu", 6, "Mayumi's", 2, "Condiments", "24 - 250 ml bottles", 15.5, 39, 0, 5, False
    .AddValues 44, "Gula Malacca", 20, "Leka Trading", 2, "Condiments", "20 - 2 kg bags", 19.45, 27, 0, 15, False
    .AddValues 61, "Sirop d'�rable", 29, "Forêts d'�rables", 2, "Condiments", "24 - 500 ml bottles", 28.5, 113, 0, 25, False
    .AddValues 63, "Vegie-spread", 7, "Pavlova, Ltd.", 2, "Condiments", "15 - 625 g jars", 43.9, 24, 0, 5, False
    .AddValues 65, "Louisiana Fiery Hot Pepper Sauce", 2, "New Orleans Cajun Delights", 2, "Condiments", "32 - 8 oz bottles", 21.05, 76, 0, 0, False
    .AddValues 66, "Louisiana Hot Spiced Okra", 2, "New Orleans Cajun Delights", 2, "Condiments", "24 - 8 oz jars", 17, 4, 100, 20, False
    .AddValues 77, "Original Frankfurter grüne Soße", 12, "Plutzer Lebensmittelgroßmärkte AG", 2, "Condiments", "12 boxes", 13, 32, 0, 15, False
    .AddValues 16, "Pavlova", 7, "Pavlova, Ltd.", 3, "Confections", "32 - 500 g boxes", 17.45, 29, 0, 10, False
    .AddValues 19, "Teatime Chocolate Biscuits", 8, "Specialty Biscuits, Ltd.", 3, "Confections", "10 boxes x 12 pieces", 9.2, 25, 0, 5, False
    .AddValues 20, "Sir Rodney's Marmalade", 8, "Specialty Biscuits, Ltd.", 3, "Confections", "30 gift boxes", 81, 40, 0, 0, False
    .AddValues 21, "Sir Rodney's Scones", 8, "Specialty Biscuits, Ltd.", 3, "Confections", "24 pkgs. x 4 pieces", 10, 3, 40, 5, False
    .AddValues 25, "NuNuCa Nuß-Nougat-Creme", 11, "Heli Süßwaren GmbH & Co. KG", 3, "Confections", "20 - 450 g glasses", 14, 76, 0, 30, False
    .AddValues 26, "Gumbär Gummibärchen", 11, "Heli Süßwaren GmbH & Co. KG", 3, "Confections", "100 - 250 g bags", 31.23, 15, 0, 0, False
    .AddValues 27, "Schoggi Schokolade", 11, "Heli Süßwaren GmbH & Co. KG", 3, "Confections", "100 - 100 g pieces", 43.9, 49, 0, 30, False
    .AddValues 47, "Zaanse koeken", 22, "Zaanse Snoepfabriek", 3, "Confections", "10 - 4 oz boxes", 9.5, 36, 0, 0, False
    .AddValues 48, "Chocolade", 22, "Zaanse Snoepfabriek", 3, "Confections", "10 pkgs.", 12.75, 15, 70, 25, False
    .AddValues 49, "Maxilaku", 23, "Karkki Oy", 3, "Confections", "24 - 50 g pkgs.", 20, 10, 60, 15, False
    .AddValues 50, "Valkoinen suklaa", 23, "Karkki Oy", 3, "Confections", "12 - 100 g bars", 16.25, 65, 0, 30, False
    .AddValues 62, "Tarte au sucre", 29, "Forêts d'�rables", 3, "Confections", "48 pies", 49.3, 17, 0, 0, False
    .AddValues 68, "Scottish Longbreads", 8, "Specialty Biscuits, Ltd.", 3, "Confections", "10 boxes x 8 pieces", 12.5, 6, 10, 15, False
    .AddValues 11, "Queso Cabrales", 5, "Cooperativa de Quesos 'Las Cabras'", 4, "Dairy Products", "1 kg pkg.", 21, 22, 30, 30, False
    .AddValues 12, "Queso Manchego La Pastora", 5, "Cooperativa de Quesos 'Las Cabras'", 4, "Dairy Products", "10 - 500 g pkgs.", 38, 86, 0, 0, False
    .AddValues 31, "Gorgonzola Telino", 14, "Formaggi Fortini s.r.l.", 4, "Dairy Products", "12 - 100 g pkgs", 12.5, 0, 70, 20, False
    .AddValues 32, "Mascarpone Fabioli", 14, "Formaggi Fortini s.r.l.", 4, "Dairy Products", "24 - 200 g pkgs.", 32, 9, 40, 25, False
    .AddValues 33, "Geitost", 15, "Norske Meierier", 4, "Dairy Products", "500 g", 2.5, 112, 0, 20, False
    .AddValues 59, "Raclette Courdavault", 28, "Gai pâturage", 4, "Dairy Products", "5 kg pkg.", 55, 79, 0, 0, False
    .AddValues 60, "Camembert Pierrot", 28, "Gai pâturage", 4, "Dairy Products", "15 - 300 g rounds", 34, 19, 0, 0, False
    .AddValues 69, "Gudbrandsdalsost", 15, "Norske Meierier", 4, "Dairy Products", "10 kg pkg.", 36, 26, 0, 15, False
    .AddValues 71, "Fløtemysost", 15, "Norske Meierier", 4, "Dairy Products", "10 - 500 g pkgs.", 21.5, 26, 0, 0, False
    .AddValues 72, "Mozzarella di Giovanni", 14, "Formaggi Fortini s.r.l.", 4, "Dairy Products", "24 - 200 g pkgs.", 34.8, 14, 0, 0, False
    .AddValues 22, "Gustaf's Knäckebröd", 9, "PB Knäckebröd AB", 5, "Grains/Cereals", "24 - 500 g pkgs.", 21, 104, 0, 25, False
    .AddValues 23, "Tunnbröd", 9, "PB Knäckebröd AB", 5, "Grains/Cereals", "12 - 250 g pkgs.", 9, 61, 0, 25, False
    .AddValues 42, "Singaporean Hokkien Fried Mee", 20, "Leka Trading", 5, "Grains/Cereals", "32 - 1 kg pkgs.", 14, 26, 0, 0, True
    .AddValues 52, "Filo Mix", 24, "G'day, Mate", 5, "Grains/Cereals", "16 - 2 kg boxes", 7, 38, 0, 25, False
    .AddValues 56, "Gnocchi di nonna Alice", 26, "Pasta Buttini s.r.l.", 5, "Grains/Cereals", "24 - 250 g pkgs.", 38, 21, 10, 30, False
    .AddValues 57, "Ravioli Angelo", 26, "Pasta Buttini s.r.l.", 5, "Grains/Cereals", "24 - 250 g pkgs.", 19.5, 36, 0, 20, False
    .AddValues 64, "Wimmers gute Semmelknödel", 12, "Plutzer Lebensmittelgroßmärkte AG", 5, "Grains/Cereals", "20 bags x 4 pieces", 33.25, 22, 80, 30, False
    .AddValues 9, "Mishi Kobe Niku", 4, "Tokyo Traders", 6, "Meat/Poultry", "18 - 500 g pkgs.", 97, 29, 0, 0, True
    .AddValues 17, "Alice Mutton", 7, "Pavlova, Ltd.", 6, "Meat/Poultry", "20 - 1 kg tins", 39, 0, 0, 0, True
    .AddValues 29, "Thüringer Rostbratwurst", 12, "Plutzer Lebensmittelgroßmärkte AG", 6, "Meat/Poultry", "50 bags x 30 sausgs.", 123.79, 0, 0, 0, True
    .AddValues 53, "Perth Pasties", 24, "G'day, Mate", 6, "Meat/Poultry", "48 pieces", 32.8, 0, 0, 0, True
    .AddValues 54, "Tourtière", 25, "Ma Maison", 6, "Meat/Poultry", "16 pies", 7.45, 21, 0, 10, False
    .AddValues 55, "Pât� chinois", 25, "Ma Maison", 6, "Meat/Poultry", "24 boxes x 2 pies", 24, 115, 0, 20, False
    .AddValues 7, "Uncle Bob's Organic Dried Pears", 3, "Grandma Kelly's Homestead", 7, "Produce", "12 - 1 lb pkgs.", 30, 15, 0, 10, False
    .AddValues 14, "Tofu", 6, "Mayumi's", 7, "Produce", "40 - 100 g pkgs.", 23.25, 35, 0, 0, False
    .AddValues 28, "Rössle Sauerkraut", 12, "Plutzer Lebensmittelgroßmärkte AG", 7, "Produce", "25 - 825 g cans", 45.6, 26, 0, 0, True
    .AddValues 51, "Manjimup Dried Apples", 24, "G'day, Mate", 7, "Produce", "50 - 300 g pkgs.", 53, 20, 0, 10, False
    .AddValues 74, "Longlife Tofu", 4, "Tokyo Traders", 7, "Produce", "5 kg pkg.", 10, 4, 20, 5, False
    .AddValues 10, "Ikura", 4, "Tokyo Traders", 8, "Seafood", "12 - 200 ml jars", 31, 31, 0, 0, False
    .AddValues 13, "Konbu", 6, "Mayumi's", 8, "Seafood", "2 kg box", 6, 24, 0, 5, False
    .AddValues 18, "Carnarvon Tigers", 7, "Pavlova, Ltd.", 8, "Seafood", "16 kg pkg.", 62.5, 42, 0, 0, False
    .AddValues 30, "Nord-Ost Matjeshering", 13, "Nord-Ost-Fisch Handelsgesellschaft mbH", 8, "Seafood", "10 - 200 g glasses", 25.89, 10, 0, 15, False
    .AddValues 36, "Inlagd Sill", 17, "Svensk Sjöföda AB", 8, "Seafood", "24 - 250 g  jars", 19, 112, 0, 20, False
    .AddValues 37, "Gravad lax", 17, "Svensk Sjöföda AB", 8, "Seafood", "12 - 500 g pkgs.", 26, 11, 50, 25, False
    .AddValues 40, "Boston Crab Meat", 19, "New England Seafood Cannery", 8, "Seafood", "24 - 4 oz tins", 18.4, 123, 0, 30, False
    .AddValues 41, "Jack's New England Clam Chowder", 19, "New England Seafood Cannery", 8, "Seafood", "12 - 12 oz cans", 9.65, 85, 0, 10, False
    .AddValues 45, "Røgede sild", 21, "Lyngbysild", 8, "Seafood", "1k pkg.", 9.5, 5, 70, 15, False
    .AddValues 46, "Spegesild", 21, "Lyngbysild", 8, "Seafood", "4 - 450 g glasses", 12, 95, 0, 0, False
    .AddValues 58, "Escargots de Bourgogne", 27, "Escargots Nouveaux", 8, "Seafood", "24 pieces", 13.25, 62, 0, 20, False
    .AddValues 73, "Röd Kaviar", 17, "Svensk Sjöföda AB", 8, "Seafood", "24 - 150 g jars", 15, 101, 0, 5, False
  End With
  
  CreateProductsList = True
  Exit Function
CreateProductsList_Err:
  OutputLn "Error creating catalog list: " & Err.Description
  Set mlstProducts = Nothing
End Function

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestProductListToJson()
  Const LOCAL_ERR_CTX As String = "TestProductListToJson"
  On Error GoTo TestProductListToJson_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Convert CRow/CList object hierarchy to JSON string"
  
  If Not CreateProductsList() Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Failed to create test suppliers list"
    GoTo TestProductListToJson_Exit
  End If
  Test_Comment mlstProducts.Count & " products."
  
  Dim oConv   As New CJsonConverter
  Dim sJson   As String
  sJson = oConv.ConvertToJson(mlstProducts)
  Test_ValueNotEqual 0, Len(sJson), "JSON conversion error"
  Test_Value "[{""ProductID"":1,""ProductName"":""Chai""", Left$(sJson, 36), "JSON unexpected conversion"
  Test_Value """ReorderLevel"":5,""Discontinued"":false}]", Right$(sJson, 39), "JSON unexpected conversion"
  
TestProductListToJson_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestProductListToJson_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestProductListToJson_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestListGroupBy1()
  Const LOCAL_ERR_CTX As String = "TestListGroupBy1"
  On Error GoTo TestListGroupBy1_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test CList GroupBy() member function"
  
  Dim lstData   As New CList

  Test_Comment "--- Testing group by with 4 elements, 1 col sublist"
  lstData.ArrayDefine Array("Name", "Age"), Array(vbString, vbString)
  
  lstData.AddValues "Barley", 8
  lstData.AddValues "Boots", 4
  lstData.AddValues "Whisker", 1
  lstData.AddValues "Daisy", 4

  Dim lstGrouped As CList
  
  Set lstGrouped = lstData.GroupBy(Array("Age"), Array("Name", "Age"))
  ListDump lstGrouped, "Grouped list"
  Test_Value 3&, lstGrouped.Count, "Group by seems incorrect, should have 3 master elements"
  
  Dim i       As Integer
  For i = 1 To lstGrouped.Count
    Test_Comment "Age: " & lstGrouped("Age", i)
    Select Case lstGrouped("Age", i)
    Case 1
      ListDump lstGrouped("__tuples", i), "Age=1"
      Test_Value 1&, lstGrouped("__tuples", i).Count, "Incorrect sublist element count for age 1"
    Case 4
      ListDump lstGrouped("__tuples", i), "Age=4"
      Test_Value 2&, lstGrouped("__tuples", i).Count, "Incorrect sublist element count for age 4"
    Case 8
      ListDump lstGrouped("__tuples", i), "Age=8"
      Test_Value 1&, lstGrouped("__tuples", i).Count, "Incorrect sublist element count for age 8"
    End Select
  Next i

TestListGroupBy1_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestListGroupBy1_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListGroupBy1_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestListGroupByWithOneElement()
  Const LOCAL_ERR_CTX As String = "TestListGroupByWithOneElement"
  On Error GoTo TestListGroupByWithOneElement_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test CList GroupBy() member function - 1 row"
  
  Dim lstData   As New CList

  Test_Comment "--- Testing group by with 1 element"
  lstData.ArrayDefine Array("Name", "Age"), Array(vbString, vbString)
  
  lstData.AddValues "Barley", 8

  Dim lstGrouped As CList
  
  Set lstGrouped = lstData.GroupBy(Array("Age"))
  Test_Value 1&, lstGrouped("__tuples", 1).Count, "Element should have tuples list of one element"
  
  Dim i       As Integer
  
  For i = 1 To lstGrouped.Count
    Test_Comment "Age: " & lstGrouped("Age", i)
    ListDump lstGrouped("__tuples", i)
    Test_Value 8, lstGrouped("__tuples", 1)("Age", 1), "Tuples list element age should be 8"
  Next i

TestListGroupByWithOneElement_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestListGroupByWithOneElement_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListGroupByWithOneElement_Exit
End Sub

#If TWINBASIC Then
[ TestCase ]
#End If
Public Sub TestListGroupByProducts()
  Const LOCAL_ERR_CTX As String = "TestListGroupByProducts"
  On Error GoTo TestListGroupByProducts_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test CList GroupBy() member function - By products"
  
  'Create a products list
  If Not CreateProductsList() Then
    Test_SetSuccess LOCAL_ERR_CTX, False, "Couldn't create product list"
    GoTo TestListGroupByProducts_Exit
  End If
  
  Dim lstGrouped As CList
  Set lstGrouped = mlstProducts.SortAndGroupBy("SupplierID")
  ListDump lstGrouped, pfDeepDump:=True


TestListGroupByProducts_Exit:
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub

TestListGroupByProducts_Err:
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestListGroupByProducts_Exit
End Sub

Public Sub TestJsonTestFiles()
  Const LOCAL_ERR_CTX As String = "TestJsonTestFiles"
  On Error GoTo TestJsonTestFiles_Err
  
  Test_BeginTestProc LOCAL_ERR_CTX
  Test_Comment "Test reading JSON text and converting to CRow/CList object hierarchy"
  
  On Error GoTo TestJsonTestFiles_Err

  Dim sJson     As String
  Dim sName     As String
  Dim sFilename As String
  Dim i         As Integer
  Dim lstData   As CList
  Dim oConv     As New CJsonConverter
  Dim oJson     As Object
  Dim rowData   As CRow
  
  Const TEST_FILES_COUNT As Integer = 4
  
  For i = 1 To TEST_FILES_COUNT
    sName = "test_json_" & i & ".txt"
    sFilename = CombinePath(GetTestDataInputDirectory(), sName)
    If Not ExistFile(sFilename) Then
      Test_Comment "File [" & sFilename & "] not found"
      GoTo next_file
    End If
    Test_Comment "Processing [" & sFilename & "]"
    sJson = GetFileText(sFilename, "utf8")
    'Debug.Print "Convert JSON to Object"
    'Debug.Print "=================== JSON:" & vbCrLf & sJson
    On Error Resume Next
    Err.Clear
    Set oJson = oConv.ParseJson(sJson)
    If Err.Number = 0 Then
      'Debug.Print "Type name of parsed object: " & TypeName(oJson)
      If TypeOf oJson Is CList Then
        Set lstData = oJson
        ListDump lstData, sName & ":CList"
      Else
        If TypeOf oJson Is CRow Then
          ListDump rowData, sName & ":CRow"
        End If
      End If
    Else
      Test_Comment "Error converting #" & Err.Number & ": " & Err.Description
    End If
    Set oJson = Nothing
next_file:
  Next i
  
TestJsonTestFiles_Exit:
  Set oJson = Nothing
  Set lstData = Nothing
  Set rowData = Nothing
  Test_EndTestProc LOCAL_ERR_CTX
  Exit Sub
  
TestJsonTestFiles_Err:
  OutputLn "TestJsonTestFiles error #" & Err.Number & ": " & Err.Description
  Test_SetSuccess LOCAL_ERR_CTX, False, Err.Description
  Resume TestJsonTestFiles_Exit
End Sub

#If TWINBASIC Then
End Module
#End If
