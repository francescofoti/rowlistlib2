Attribute VB_Name = "MRunTests"
#If TWINBASIC Then
Module MRunTests
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

#If VBA7 Then
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long
#Else
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
#End If

#If VBA7 Then
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Public Const TESTPARAM_NWINDPATH    As String = "nwindpath"
Public Const TESTPARAM_TESTDATAPATH As String = "testdatapath"
Public Const TESTPARAM_ANALJCONV    As String = "analjconv"

Private mfIsRunning As Boolean

Public Function Test_IsSuiteRunning() As Boolean
  Test_IsSuiteRunning = mfIsRunning
End Function

Public Sub Test_BeginRun()
  #If CONOUTX Then
    frmMain.Show
  #End If
  mfIsRunning = True
End Sub

Public Sub Test_EndRun()
  mfIsRunning = False
End Sub

Public Sub SetupTests()
  Test_SetParam SUITEPARAM_UNATTENDED, True
  Test_SetParam SUITEPARAM_LISTDUMPOFF, True
  Test_SetParam TESTPARAM_ANALJCONV, True
  Test_SetParam SUITEPARAM_INPUTITERATIONS, False
  Test_SetParam SUITEPARAM_DEFITERCT, 100
  Test_SetParam SUITEPARAM_FAKEVIEWFILE, True
  
  'nwind db is in the parent path
  Dim sPath As String
  #If MSACCESS Then
    sPath = CurrentProject.Path
    Test_SetParam TESTPARAM_NWINDPATH, CombinePath(sPath, "..\testdata\json\input\Ms Access Northwind Database.accdb")
  #Else
    sPath = App.Path
    #If TWINBASIC = 0 Then
      'For VB path is also different
      Test_SetParam TESTPARAM_TESTDATAPATH, CombinePath(sPath, "..\..\..\testdata")
      Test_SetParam TESTPARAM_NWINDPATH, CombinePath(sPath, "..\..\..\testdata\json\input\Ms Access Northwind Database.accdb")
    #End If
  #End If
  Test_SetParam TESTPARAM_ANALJCONV, True
End Sub

Public Function AreControlKeysDown() As Boolean
  Const VK_LCTRL      As Long = &HA2  ' 162
  Const VK_RCTRL      As Long = &HA3  ' 163
  If (GetKeyState(VK_LCTRL) < 0) And (GetKeyState(VK_RCTRL) < 0) Then
    AreControlKeysDown = True
  End If
End Function

Public Function TimedPause(ByVal piSeconds As Integer, ByVal psReason As String, Optional ByVal pfCanAbort As Boolean = False) As Boolean
  Dim t  As Single
  Dim i  As Integer
  Dim sBuf  As String
  Dim sTemp As String
  
  Const MAX_LINE_CHARS As Integer = 60
  
  TimedPause = True
  
  ConOut "(Timer) " & psReason
  If pfCanAbort Then
    ConOutLn " (press left+right CTRL to abort)"
  Else
    ConOutLn ""
  End If
  
  sBuf = String$(MAX_LINE_CHARS + 5, " ")
  For i = 1 To MAX_LINE_CHARS
    If (i = 1) Or ((i Mod 10) = 0) Then
      sTemp = "¦" & i
      Mid$(sBuf, i, Len(sTemp)) = sTemp
    End If
  Next i

  ConOutLn sBuf
  sBuf = String$(MAX_LINE_CHARS, "-")
  For i = 1 To MAX_LINE_CHARS
    If (i = 1) Or ((i Mod 10) = 0) Then
      Mid$(sBuf, i, 1) = "+"
    End If
  Next i
  If piSeconds <= MAX_LINE_CHARS Then
    Mid$(sBuf, piSeconds + 1, 1) = "*"
  End If
  ConOutLn sBuf
    
  Dim xx As Integer
  Dim tick As Integer
  Dim lasttick As Integer
  t = Timer
  xx = 0: lasttick = 0
  Do While (Timer - t) < piSeconds
    tick = (CInt(Timer - t) Mod 2)
    If (lasttick <> tick) Then
      lasttick = tick
      Debug.Print ">"; 'Do not use ConOut here (it's buffered)
      xx = xx + 1
      If xx = MAX_LINE_CHARS Then
        Debug.Print " " & Format$(Time, "hh:mm:ss")
        xx = 0
      End If
    End If
    DoEvents
    If AreControlKeysDown() Then
      If pfCanAbort Then
        ConOutLn "...aborted!"
        TimedPause = False
        Exit Function
      End If
    End If
  Loop
  ConOutLn "...dring!"
End Function

Public Sub RunAllTests()
  On Error GoTo RunAllTests_Err
  
  SetupTests
  Test_BeginRun
  
  Dim i             As Long
  Dim lIterCount    As Long
  Dim sngStartTimer As Single
  
  Const ITER_WAIT_SECS As Integer = 3
  
  If Nz(Test_GetParam(SUITEPARAM_INPUTITERATIONS), True) Then
    lIterCount = LongChooseBox("Number of iterations:", "Run how many times ?", "1", 1000)
    If lIterCount = 0 Then Exit Sub
  Else
    If Nz(Test_GetParam(SUITEPARAM_DEFITERCT), 0) > 0 Then
      lIterCount = Test_GetParam(SUITEPARAM_DEFITERCT)
    Else
      lIterCount = 1
    End If
  End If
  
  For i = 1 To lIterCount
    sngStartTimer = Timer
    ConOutLn "****************************************************************"
    ConOutLn " ITERATION #" & i & "/" & lIterCount & " started at " & Now
    ConOutLn "  | | |"
    ConOutLn "  v v v"
    ConOutLn "****************************************************************"
    ConOutLn ""
  
    If Not TimedPause(ITER_WAIT_SECS, "Waiting to start MRowTests.Test_RunAllTests", True) Then Exit For
    MRowTests.Test_RunAllTests
    If Not Test_LastSuiteSuccess() Then GoTo RunAllTests_Exit
    
    If Not TimedPause(ITER_WAIT_SECS, "Waiting to start MListTests.Test_RunAllTests", True) Then Exit For
    MListTests.Test_RunAllTests
    If Not Test_LastSuiteSuccess() Then GoTo RunAllTests_Exit

    If Not TimedPause(ITER_WAIT_SECS, "Waiting to start MJsonTests.Test_RunAllTests", True) Then Exit For
    MJsonTests.Test_RunAllTests
    
    If Not Test_LastSuiteSuccess() Then GoTo RunAllTests_Exit

    ConOutLn "****************************************************************"
    ConOutLn "  ^ ^ ^"
    ConOutLn "  | | |"
    ConOutLn " ITERATION #" & i & "/" & lIterCount & " ended at " & Now
    ConOutLn " ELAPSED TIME: " & (Timer - sngStartTimer) & " secs"
    ConOutLn "****************************************************************"
    
    If i < lIterCount Then
      If Not TimedPause(ITER_WAIT_SECS, "Waiting to start next iteration", True) Then
        'cancelled
        Exit For
      End If
    End If
  
  Next i
  
RunAllTests_Exit:
  Test_EndRun
  Exit Sub

RunAllTests_Err:
  ConOutLn "RunAllTests() failed: " & Err.Description
  Resume RunAllTests_Exit
End Sub

Public Sub RunProblem()
  SetupTests
  'MJsonTests.TestDatabase_Orders
  MJsonTests.TestDatabase_Queries
End Sub

#If TWINBASIC Then
End Module
#End If

