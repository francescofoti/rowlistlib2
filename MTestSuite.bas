Attribute VB_Name = "MTestSuite"
#If TWINBASIC Then
Module MTestSuite
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Public Const APP_VERSION As String = "02.04.04"

'For the Test_ system
Private mfTestCommentsOff As Boolean

Private mlstProcs As CList
Private Const COL_PROCNAME    As String = "procname"
Private Const COL_TIMERBEGIN  As String = "timerbegin"
Private Const COL_TIMEREND    As String = "timerend"
Private Const COL_BEGINAT     As String = "beginat"
Private Const COL_ENDAT       As String = "endat"
Private Const COL_TIMEDIFF    As String = "difftime"
Private Const COL_SUCCESS     As String = "success"
Private Const COL_MESSAGE     As String = "message"
Private Const COL_CALLCT      As String = "callct"

Private mfLastSuiteSuccess As Boolean
Private mlstParams As CList

Public Const SUITEPARAM_UNATTENDED      As String = "unattended"
Public Const SUITEPARAM_LISTDUMPOFF     As String = "listdumpoff"
Public Const SUITEPARAM_INPUTITERATIONS As String = "inputiters"
Public Const SUITEPARAM_DEFITERCT       As String = "defiterct"
Public Const SUITEPARAM_FAKEVIEWFILE    As String = "fakeviewer"

Private mlErr       As Long
Private msErrDesc   As String
Private msErrCtx    As String

Private Sub ClearErr()
  mlErr = 0&
  msErrCtx = ""
  msErrDesc = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErrNo As Long, ByVal psErrDesc As String)
  msErrCtx = psErrCtx
  mlErr = plErrNo
  msErrDesc = psErrDesc
End Sub

Public Function Test_LastErr() As Long
  Test_LastErr = mlErr
End Function

Public Function Test_LastErrDesc() As String
  Test_LastErrDesc = msErrDesc
End Function

Public Function Test_LastErrCtx() As String
  Test_LastErrCtx = msErrCtx
End Function

Public Function StrTrueFalse(ByVal pfFlag As Boolean) As String
  StrTrueFalse = IIf(pfFlag, "True", "False")
End Function

Private Sub InitParamsList()
  If Not mlstParams Is Nothing Then Exit Sub
  Set mlstParams = New CList
  mlstParams.ArrayDefine Array("name", "value"), Array(vbString, vbVariant)
End Sub

Public Function Test_GetParam(ByVal psParamName As String) As Variant
  Const LOCAL_ERR_CTX As String = "Test_GetParam"
  On Error GoTo Test_GetParam_Err
  
  Test_GetParam = Null
  InitParamsList
  
  Dim iFind   As Long
  iFind = mlstParams.Find("name", psParamName)
  If iFind > 0& Then
    
    If Not mlstParams.IsItemObject("value", iFind) Then
      Test_GetParam = mlstParams("value", iFind)
    Else
      Set Test_GetParam = mlstParams("value", iFind)
    End If
  End If
  
Test_GetParam_Exit:
  Exit Function
Test_GetParam_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Test_GetParam_Exit
End Function

Public Sub Test_SetParam(ByVal psParamName As String, ByVal pvParamValue As Variant)
  Const LOCAL_ERR_CTX As String = "Test_SetParam"
  On Error GoTo Test_SetParam_Err
  
  InitParamsList
  
  Dim iFind   As Long
  iFind = mlstParams.Find("name", psParamName)
  If iFind > 0& Then
    If Not IsObject(pvParamValue) Then
      mlstParams("value", iFind) = pvParamValue
    Else
      Set mlstParams("value", iFind) = pvParamValue
    End If
  Else
    mlstParams.AddValues psParamName, pvParamValue
  End If
  
Test_SetParam_Exit:
  Exit Sub
Test_SetParam_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Test_SetParam_Exit
End Sub

Public Sub Test_ValueNotEqual(ByVal pvShouldNotBe As Variant, ByVal pvValue As Variant, ByVal psMessage As String)
  On Error Resume Next
  Dim fIsDifferent    As Boolean
  
  If Not IsNull(pvShouldNotBe) Then
    If Not IsNull(pvValue) Then
      fIsDifferent = CBool(pvShouldNotBe <> pvValue)
    Else
      fIsDifferent = True
    End If
  Else
    'if pvShouldNotBe is null, then pvValue must not be null
    fIsDifferent = Not IsNull(pvValue)
  End If
  
  If Err.Number = 0 Then
    If fIsDifferent Then
      ConOut " [Ok   ]"
    Else
      'ConOut "*[Error]"
      Err.Raise 51&, "Test_Value", "Values should be different [" & VariantAsString(pvShouldNotBe) & "]=[" & VariantAsString(pvValue) & "]"
    End If
  End If
  ConOutLn " " & psMessage
End Sub

Public Sub Test_Value(ByVal pvShouldBe As Variant, ByVal pvValue As Variant, ByVal psMessage As String)
  Dim fTestValueResult    As Boolean
  
  If Not IsObject(pvValue) Then
    If Not IsNull(pvShouldBe) Then
      If Not IsNull(pvValue) Then
        fTestValueResult = CBool(pvValue = pvShouldBe)
      End If
    Else
      'if pvShouldBe is null, then pvValue has to be null (null=null)
      fTestValueResult = IsNull(pvValue)
    End If
  Else
    fTestValueResult = CBool(pvValue Is pvShouldBe)
  End If
  
  If fTestValueResult Then
    ConOut " [Ok   ]"
  Else
    'ConOut "*[Error]"
    Err.Raise 51&, "Test_Value", "Expected value [" & VariantAsString(pvShouldBe) & "] but received [" & VariantAsString(pvValue) & "]"
  End If
  ConOutLn " " & psMessage
End Sub

Public Sub Test_ShowComments(ByVal pfShow As Boolean)
  mfTestCommentsOff = Not pfShow
End Sub

Public Sub Test_Comment(ByVal psMessage As String)
  If Not mfTestCommentsOff Then ConOutLn " [Cmmnt] " & psMessage
End Sub

Public Sub Test_Banner(ByVal psMessage As String)
  ConOutLn " [Messg] " & psMessage
End Sub

Private Sub InitProcList()
  If mlstProcs Is Nothing Then
    Set mlstProcs = New CList
    mlstProcs.ArrayDefine _
      Array(COL_PROCNAME, COL_BEGINAT, COL_ENDAT, COL_TIMERBEGIN, COL_TIMEREND, COL_TIMEDIFF, COL_SUCCESS, COL_MESSAGE, COL_CALLCT), _
      Array(vbString, vbString, vbString, vbSingle, vbSingle, vbSingle, vbBoolean, vbString, vbLong)
  End If
End Sub

Private Function GetProcIndex(ByVal psProcName As String) As Long
  GetProcIndex = mlstProcs.Find(COL_PROCNAME, psProcName)
End Function

Public Sub Test_BeginTestProc(ByVal psProcName As String)
  InitProcList
  Dim iProc         As Long
  Dim sngStartTime  As Single
  
  iProc = GetProcIndex(psProcName)
  sngStartTime = Timer
  If iProc > 0 Then
    mlstProcs(COL_TIMERBEGIN, iProc) = sngStartTime
    mlstProcs(COL_TIMEREND, iProc) = Null
    mlstProcs(COL_BEGINAT, iProc) = Format$(Now, "hh:mm:ss")
    mlstProcs(COL_ENDAT, iProc) = ""
    mlstProcs(COL_CALLCT, iProc) = Nz(mlstProcs(COL_CALLCT, iProc), 0) + 1
  Else
    mlstProcs.AddValues psProcName, Format$(Now, "hh:mm:ss"), "", sngStartTime, Null, Null, Null, Null, 1
  End If
  ConOutLn " [PROC ] [" & psProcName & "] Starting at " & Format$(Now, "hh:mm:ss")
End Sub

Public Sub Test_EndTestProc(ByVal psProcName As String)
  Dim iProc         As Long
  Dim sngEndTime    As Single
  
  iProc = GetProcIndex(psProcName)
  sngEndTime = Timer
  If iProc > 0 Then
    mlstProcs(COL_TIMEREND, iProc) = sngEndTime
    mlstProcs(COL_TIMEDIFF, iProc) = sngEndTime - mlstProcs(COL_TIMERBEGIN, iProc)
    mlstProcs(COL_ENDAT, iProc) = Format$(Now, "hh:mm:ss")
    ConOutLn " [/PROC] [" & psProcName & "] Ended at " & Format$(Now, "hh:mm:ss") & " time:" & mlstProcs(COL_TIMEDIFF, iProc) & " ms"
  Else
    ConOutLn " [/PROC] [" & psProcName & "] Ended at " & Format$(Now, "hh:mm:ss")
  End If
End Sub

Public Sub Test_SetSuccess(ByVal psProcName As String, ByVal pfSuccess As Boolean, Optional ByVal pvMessage As Variant = Null)
  ConOutLn " [RESLT] [" & psProcName & "] :" & IIf(pfSuccess, "OK", "FAIL")
  Dim iProc         As Long
  iProc = GetProcIndex(psProcName)
  If iProc > 0 Then
    mlstProcs(COL_SUCCESS, iProc) = pfSuccess
    mlstProcs(COL_MESSAGE, iProc) = pvMessage
  End If
End Sub

Public Sub Test_BeginSuite(ByVal psMessage As String)
  ConOutLn String$(50, "=")
  ConOutLn ""
  ConOutLn " [SUITE] " & psMessage
  ConOutLn ""
  ConOutLn String$(50, "=")
  mfLastSuiteSuccess = False
End Sub

Public Function Test_LastSuiteSuccess() As Boolean
  Test_LastSuiteSuccess = mfLastSuiteSuccess
End Function

Public Sub Test_EndSuite()
  Dim i           As Long
  Dim lErrCt      As Long

  If Not mlstProcs Is Nothing Then
    ListDump mlstProcs, "Timings", COL_PROCNAME & ":40;" & COL_TIMERBEGIN & ":10;" & _
                        COL_TIMEREND & ":10;" & COL_CALLCT & ":8;" & _
                        COL_BEGINAT & ":10;" & COL_ENDAT & ":10", _
                        pfIgnoreUnattended:=True
    
    If mlstProcs.Count > 0 Then
      For i = 1 To mlstProcs.Count
        If Not Nz(mlstProcs(COL_SUCCESS, i), True) Then
          lErrCt = lErrCt + 1&
          ConOutLn " [SUITE] Error #" & lErrCt & " : " & mlstProcs(COL_MESSAGE, i)
        End If
      Next i
    End If
  End If
  
  ConOutLn String$(50, "=")
  ConOutLn ""
  ConOutLn " [SUITE] (finished) " & lErrCt & " error(s)"
  ConOutLn ""
  ConOutLn String$(50, "=")
  If lErrCt = 0& Then
    mfLastSuiteSuccess = True
    ConOutLn " [SUITE] SUCCEEDED"
  Else
    mfLastSuiteSuccess = False
    ConOutLn " [SUITE] FAILED"
  End If
  ConOutLn String$(50, "=")
  
  Set mlstProcs = Nothing
End Sub

Private Function SliceToArray(ByRef pasRetArray() As String, ByVal psText As String, ByVal piColWidth As Integer) As Integer
  Dim iSize     As Integer
  Dim iTextLen  As Integer
  Dim i         As Integer
  
  iTextLen = Len(psText)
  If iTextLen = 0 Then Exit Function
  
  iSize = iTextLen \ piColWidth
  If (iSize * piColWidth) < iTextLen Then
    iSize = iSize + 1
  End If
  ReDim pasRetArray(1 To iSize) As String
  
  For i = 1 To iSize
    pasRetArray(i) = Mid$(psText, (i - 1) * piColWidth + 1, piColWidth)
  Next i
  
  SliceToArray = iSize
End Function

Public Sub SideBySideDebugPrint( _
  ByVal ps1 As String, ps2 As String, _
  Optional ByVal piColChars As Integer = 40, _
  Optional ByVal pfShowFirstDiffOnly As Boolean = True, _
  Optional ByVal psTitle As String = "")
  Dim iLen1 As Integer
  Dim iLen2 As Integer
  Dim i1    As Integer
  Dim i2    As Integer
  Dim iLine As Integer
  Dim sLine As String
  Dim as1() As String
  Dim as2() As String
  Dim i1Count As Integer
  Dim i2Count As Integer
  Dim sLine1  As String
  Dim sLine2  As String
  Dim iFmtNumberLen As Integer
  
  Dim fShowDiff As Boolean
  Dim fDifferent As Boolean
  
  If Len(psTitle) > 0 Then
    ConOutLn psTitle
  End If
  
  iLine = 1
  i1 = Len(ps1)
  i2 = Len(ps2)
  
  i1Count = SliceToArray(as1, ps1, piColChars)
  i2Count = SliceToArray(as2, ps2, piColChars)
  
  iFmtNumberLen = 3
  iLine = 1
  fShowDiff = True
  Do While (iLine <= i1Count) Or (iLine <= i2Count)
    If iLine <= i1Count Then
      sLine1 = PadRight(as1(iLine), piColChars, " ")
    Else
      sLine1 = Space$(piColChars)
    End If
    If iLine <= i2Count Then
      sLine2 = PadRight(as2(iLine), piColChars, " ")
    Else
      sLine2 = Space$(piColChars)
    End If
    
    ConOut Format$(iLine, String$(iFmtNumberLen, "0")) & ":"
    fDifferent = CBool(sLine1 <> sLine2)
    If fDifferent Then
      ConOut "*"
    Else
      ConOut " "
    End If
    ConOutLn sLine1 & " | " & sLine2
    If fDifferent And fShowDiff Then
      ConOut Space$(iFmtNumberLen + 2)
      For i1 = 1 To Len(sLine1)
        If Mid$(sLine1, i1, 1) = Mid$(sLine2, i1, 1) Then
          ConOut " "
        Else
          ConOut "^"
        End If
      Next i1
      ConOut " | "
      For i2 = 1 To Len(sLine2)
        If Mid$(sLine1, i2, 1) = Mid$(sLine2, i2, 1) Then
          ConOut " "
        Else
          ConOut "^"
        End If
      Next i2
      ConOut ""
      
      ConOut Space$(iFmtNumberLen + 2)
      For i1 = 1 To Len(sLine1)
        ConOut CStr(i1 Mod 10)
      Next i1
      ConOut " | "
      For i2 = 1 To Len(sLine2)
        ConOut CStr(i2 Mod 10)
      Next i2
      ConOut ""
      
      If pfShowFirstDiffOnly Then
        fShowDiff = False
      End If
    End If
    iLine = iLine + 1
  Loop
End Sub

Public Function CurrentProjectPath() As String
  #If MSACCESS Then
    CurrentProjectPath = NormalizePath(CurrentProject.Path)
  #Else
    CurrentProjectPath = NormalizePath(StripFileName(App.Path))
  #End If
End Function

#If TWINBASIC Then
End Module
#End If
