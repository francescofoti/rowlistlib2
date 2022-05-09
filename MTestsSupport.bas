Attribute VB_Name = "MTestsSupport"
#If TWINBASIC Then
Module MTestsSupport
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

#If Win64 Then
  'Win64 API
  Type UINT64
      LowPart As Long
      HighPart As Long
  End Type
  Private Const BSHIFT_32 = 4294967296# ' 2 ^ 32
  Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As Long
  Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As Long

  'Needed for chronometer functions
  Private mcurFrequency   As UINT64
  Private mcurChronoStart As UINT64

#Else
  'Win32 API
  Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
  Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

  'Needed for chronometer functions
  Private mcurFrequency   As Currency
  Private mcurChronoStart As Currency

#End If

'Default names for test data and log files sub directories.
'Assumed to be subdirectories of App.Path by default.
'Subdirectories are created if not existing, when needed.
Private Const DEFAULT_ROOTDIR_TESTDATA  As String = "testdata"
Private Const DEFAULT_ROOTDIR_LOGS      As String = "log"
'Pathes will be solved from App.Path by default, if no specific
'directory is specified in these two masRootDir().
Public Enum eManagedFileGroup
  eFileGroupTestData
  eFileGroupLogs
End Enum
Private masRootDir(eManagedFileGroup.eFileGroupTestData To eManagedFileGroup.eFileGroupLogs) As String

'ConOut/Ln need a temporary line buffer
Private msConOutLineBuffer As String

'Output file
Private msTraceFilename As String
Private mhOutput        As Integer

Public Function GetRootDir(ByVal peManagedFileGroup As eManagedFileGroup) As String
  If (peManagedFileGroup < eManagedFileGroup.eFileGroupTestData) Or _
     (peManagedFileGroup > eManagedFileGroup.eFileGroupLogs) Then
    Exit Function
  End If
  
  Dim sRootDir    As String
  
  sRootDir = masRootDir(peManagedFileGroup)
  If Len(sRootDir) = 0 Then
    #If MSACCESS Then
      sRootDir = CurrentProject.Path
    #Else
      sRootDir = App.Path
    #End If
    Select Case peManagedFileGroup
      Case eManagedFileGroup.eFileGroupTestData
        sRootDir = CombinePath(sRootDir, DEFAULT_ROOTDIR_TESTDATA)
      Case eManagedFileGroup.eFileGroupLogs
        sRootDir = CombinePath(sRootDir, DEFAULT_ROOTDIR_LOGS)
    End Select
  End If
  
  GetRootDir = sRootDir
End Function

Public Sub SetRootDir( _
  ByVal peManagedFileGroup As eManagedFileGroup, _
  ByVal psNewRootDir As String)
  If (peManagedFileGroup < eManagedFileGroup.eFileGroupTestData) Or _
    (peManagedFileGroup > eManagedFileGroup.eFileGroupLogs) Then
    Exit Sub
  End If
  masRootDir(peManagedFileGroup) = psNewRootDir
End Sub

'
' Support functions
'

'Get a random long
Public Function GetRandom(ByVal iLo As Long, ByVal iHi As Long) As Long
  GetRandom = Int(iLo + (Rnd * (iHi - iLo + 1)))
End Function

'
' A precise chronometer for accurate timing
'

#If Win64 Then
Public Function U64Dbl(U64 As UINT64) As Double
    Dim lDbl As Double, hDbl As Double
    lDbl = U64.LowPart
    hDbl = U64.HighPart
    If lDbl < 0 Then lDbl = lDbl + BSHIFT_32
    If hDbl < 0 Then hDbl = hDbl + BSHIFT_32
    U64Dbl = lDbl + BSHIFT_32 * hDbl
End Function
#End If

Public Sub ChronoStart()
  #If Win64 Then
    If (mcurFrequency.HighPart = 0) And (mcurFrequency.LowPart = 0) Then QueryPerformanceFrequency mcurFrequency
  #Else
    If mcurFrequency = 0 Then QueryPerformanceFrequency mcurFrequency
  #End If
  QueryPerformanceCounter mcurChronoStart
End Sub

Public Function ChronoTime() As String
#If Win64 Then
  Dim curFrequency As UINT64
  Dim dblElapsed   As Double
  QueryPerformanceCounter curFrequency
  If (mcurFrequency.LowPart = 0) And (mcurFrequency.HighPart = 0) Then
    curFrequency.LowPart = 0
    curFrequency.HighPart = 0
  Else
    dblElapsed = (U64Dbl(curFrequency) - U64Dbl(mcurChronoStart)) / U64Dbl(mcurFrequency)
  End If
  ChronoTime = CStr(dblElapsed)
#Else
  Dim curFrequency As Currency
  QueryPerformanceCounter curFrequency
  If mcurFrequency = 0 Then
    curFrequency = 0
  Else
    curFrequency = (curFrequency - mcurChronoStart) / mcurFrequency
  End If
  ChronoTime = CStr(curFrequency)
#End If
End Function

' Debug output in file and the debug window
' ConOut/Ln : output only on console device
' Output/Ln : output in log file (if open), and also on console device

Public Sub ConOut(ByVal psText As String)
  msConOutLineBuffer = msConOutLineBuffer & psText
  #If CONOUTX Then
    frmMain.ConOut psText
  #End If
End Sub

Public Sub ConOutLn(ByVal psText As String)
  If Len(msConOutLineBuffer) > 0 Then
    Debug.Print msConOutLineBuffer;
    msConOutLineBuffer = ""
  End If
  Debug.Print psText
  #If CONOUTX Then
    On Error Resume Next
    frmMain.Show
    frmMain.ConOutLn psText
  #End If
End Sub

Public Sub OutputLn(Optional ByRef sOutput As String = "")
  If mhOutput Then
    Print #mhOutput, sOutput
  End If
  #If CONOUTX Then
    frmMain.Show
  #End If
  ConOutLn sOutput
End Sub

Public Sub Output(ByRef sOutput As String)
  If mhOutput Then
    Print #mhOutput, sOutput;
  End If
  ConOut sOutput
End Sub

Public Sub OutputBanner(ByVal sSubName As String, ByVal sDescr As String)
  Dim sBanner   As String
  
  sBanner = vbCrLf & String$(60, "=") & vbCrLf & sSubName & vbCrLf & sDescr & vbCrLf & String$(60, "=") & vbCrLf
  OutputLn sBanner
  OutputLn "Running " & sSubName & vbCrLf & vbCrLf & sDescr
End Sub

Public Function SolveFileName(ByVal peManagedFileGroup As eManagedFileGroup, ByVal psFilename As String) As String
  SolveFileName = CombinePath(GetRootDir(peManagedFileGroup), psFilename)
End Function

Public Function OpenTraceOutputFile( _
    Optional ByVal psFilenameOnly As String = "", _
    Optional ByVal pfOverWrite As Boolean = False _
  ) As Boolean
  Dim fIsOpen     As Boolean
  Dim sOutputFile As String
  Dim sMsg        As String
  Dim sChoice     As String
  Dim sDir        As String
  
  On Error GoTo OpenTraceOutputFile_Err
  
  sDir = GetRootDir(eManagedFileGroup.eFileGroupLogs)
  If Not ExistDir(sDir) Then
    MkDir sDir
  End If
  
  If Len(psFilenameOnly) > 0 Then
    sOutputFile = SolveFileName(eManagedFileGroup.eFileGroupLogs, psFilenameOnly)
  Else
    sOutputFile = SolveFileName(eManagedFileGroup.eFileGroupLogs, "output.txt")
  End If
  mhOutput = FreeFile
  If Not ExistFile(sOutputFile) Then
    Open sOutputFile For Output As #mhOutput
  Else
    If Not pfOverWrite Then
      Open sOutputFile For Append As #mhOutput
    Else
      Open sOutputFile For Output As #mhOutput
    End If
  End If
  msTraceFilename = sOutputFile
  fIsOpen = True
  
  OpenTraceOutputFile = True
  Exit Function
  
OpenTraceOutputFile_Err:
  If Not Nz(Test_GetParam(SUITEPARAM_UNATTENDED), False) Then
    MsgBox Err.Number & ": " & Err.Description, vbCritical
  Else
    Debug.Print "OpenTraceOutputFile failed: " & Err.Number & ": " & Err.Description
  End If
  If fIsOpen Then
    Close #mhOutput
  End If
  mhOutput = 0
End Function

Public Sub CloseTraceOutputFile()
  If mhOutput <> 0 Then
    On Error Resume Next
    Close #mhOutput
    mhOutput = 0
  End If
  msTraceFilename = ""
End Sub

Public Function FindFile( _
    ByVal psFilename As String, _
    ByRef psRetFilePath As String, _
    ByRef pfRetIsInPATH As Boolean, _
    Optional ByVal pavAdditionalPathes As Variant = Null _
  ) As Boolean
  'First look into current project path, or in \bin subdirectory
  Dim sLookupPath     As String
  Dim sFilename       As String
  Dim fOK             As Boolean
  Dim fInPath         As Boolean
  
  On Error GoTo FindFile_Err
  psRetFilePath = ""
  pfRetIsInPATH = False
  
  'Look into project path
  #If MSACCESS Then
    sLookupPath = CurrentProject.Path
  #Else
    sLookupPath = App.Path
  #End If
  sFilename = CombinePath(sLookupPath, psFilename)
  fOK = ExistFile(sFilename)
  If Not fOK Then
    'Look into \bin subdir
    sLookupPath = CombinePath(sLookupPath, "bin")
    sFilename = CombinePath(sLookupPath, psFilename)
    fOK = ExistFile(sFilename)
    If Not fOK Then
      Dim iPathCt     As Integer
      Dim asPath()    As String
      Dim i           As Integer
      iPathCt = SplitString(asPath(), Environ$("PATH"), ";")
      For i = 1 To iPathCt
        sLookupPath = asPath(i)
        sFilename = CombinePath(sLookupPath, psFilename)
        fOK = ExistFile(sFilename)
        If fOK Then
          fInPath = True
          Exit For
        End If
      Next i
      
      'Search in additional pathes (if provided)
      If Not fOK Then
        If Not IsNull(pavAdditionalPathes) Then
          If IsArray(pavAdditionalPathes) Then
            On Error Resume Next
            iPathCt = UBound(pavAdditionalPathes) - LBound(pavAdditionalPathes) + 1
            If (Err.Number = 0) And (iPathCt > 0) Then
              On Error GoTo FindFile_Err
              For i = LBound(pavAdditionalPathes) To UBound(pavAdditionalPathes)
                sLookupPath = pavAdditionalPathes(i)
                sFilename = CombinePath(sLookupPath, psFilename)
                fOK = ExistFile(sFilename)
                If fOK Then
                  fInPath = False
                  Exit For
                End If
              Next i
            End If
          End If
        End If
      End If
    End If
  End If

  If fOK Then
    psRetFilePath = sLookupPath
    pfRetIsInPATH = fInPath
  End If
  
FindFile_Exit:
  FindFile = fOK
  Exit Function

FindFile_Err:
  OutputLn "FindFile error #" & Err.Number & ": " & Err.Description
  Resume FindFile_Exit
End Function

Public Function TraceEditorEXE() As String
  Dim sEditorEXE  As String
  Dim sEditorPath As String
  Dim fIsInPath   As Boolean
  Dim fOK         As Boolean
  Dim avAdditionalPath As Variant
  
  avAdditionalPath = Array(CombinePath(Environ$("ProgramFiles"), "Notepad++"), CombinePath(Environ$("ProgramFiles(x86)"), "Notepad++"))
  sEditorEXE = "notepad++.exe"
  fOK = FindFile(sEditorEXE, sEditorPath, fIsInPath, avAdditionalPath)
  If fOK Then
    sEditorEXE = CombinePath(sEditorPath, sEditorEXE)
  Else
    sEditorEXE = "notepad.exe"
  End If
  
  TraceEditorEXE = sEditorEXE
End Function

Public Sub ViewTraceOutputFile(Optional ByVal psFilenameOnly As String = "")
  Dim sOutputFile As String
  On Error Resume Next
  If Len(psFilenameOnly) > 0 Then
    sOutputFile = SolveFileName(eManagedFileGroup.eFileGroupLogs, psFilenameOnly)
  Else
    sOutputFile = SolveFileName(eManagedFileGroup.eFileGroupLogs, "output.txt")
  End If
  If ExistFile(sOutputFile) Then
    If Not Nz(Test_GetParam(SUITEPARAM_UNATTENDED), False) Then
      Shell TraceEditorEXE() & " " & sOutputFile, vbMaximizedFocus
    Else
      Test_Comment "[blocked] open file [" & sOutputFile & "]"
    End If
  Else
    Debug.Print "Trace file [" & sOutputFile & "] doesn't exist", vbCritical
  End If
End Sub

Private Function GetTestDataRootDirectory() As String
  Dim sPath As String
  sPath = Test_GetParam(TESTPARAM_TESTDATAPATH) & ""
  If Len(sPath) = 0 Then
    GetTestDataRootDirectory = GetRootDir(eManagedFileGroup.eFileGroupTestData)
  Else
    GetTestDataRootDirectory = sPath
  End If
End Function

Public Function GetTestDataInputDirectory() As String
  GetTestDataInputDirectory = CombinePath(GetTestDataRootDirectory(), "json\input")
End Function

Public Function GetTestDataOutputDirectory() As String
  GetTestDataOutputDirectory = CombinePath(GetTestDataRootDirectory(), "json\output")
End Function

'Cut string before trailing chr$(0)
Public Function CtoVB(ByRef pszString As String) As String
  Dim i   As Long
  i = InStr(pszString, Chr$(0))
  If i Then
    CtoVB = Left$(pszString, i - 1&)
  Else
    CtoVB = pszString
  End If
End Function

Public Function PadRight(ByVal psText As String, ByVal piLen As Integer, ByVal psChar As String) As String
  Dim iLen As Integer
  iLen = Len(psText)
  If iLen < piLen Then
    PadRight = psText & String$(piLen - iLen, Asc(psChar))
  Else
    PadRight = psText
  End If
End Function

Public Function PadLeft(ByVal psText As String, ByVal piLen As Integer, ByVal psChar As String) As String
  Dim iLen As Integer
  iLen = Len(psText)
  If iLen < piLen Then
    PadLeft = String$(piLen - iLen, Asc(psChar)) & psText
  Else
    PadLeft = psText
  End If
End Function

Public Function LTrimChar(ByVal psStrip As String, ByVal psTrimChar As String) As String
  If Len(psStrip) Then
    While Left$(psStrip, 1&) = psTrimChar
      psStrip = Right$(psStrip, Len(psStrip) - 1&)
    Wend
  End If
  LTrimChar = psStrip
End Function

Public Function Max(ByVal V1 As Variant, ByVal V2 As Variant) As Variant
  If V1 > V2 Then
    Max = V1
  Else
    Max = V2
  End If
End Function
  
Public Function Min(ByVal V1 As Variant, ByVal V2 As Variant) As Variant
  If V1 < V2 Then
    Min = V1
  Else
    Min = V2
  End If
End Function
    
'max 9 milions (9 999 999)
Public Function LongChooseBox(ByVal sText As String, ByVal sTitle As String, ByVal sDefault As String, ByVal lMax As Long, Optional ByVal lMin As Long = 1&) As Long
  Dim sInput    As String
  Dim fValid    As Boolean
  Dim lRet      As Long
  
  If lMin < 1& Then lMin = 1&
  If lMax < lMin Then
    lMax = lMin 'sounds dummy...
  End If
  
  Do
    sInput = InputBox$(sText, sTitle, sDefault)
    If Len(sInput) Then
      If IsNumeric(sInput) Then
        If Len(sInput) <= 7 Then
          lRet = CLng(Val(sInput))
          If (lRet >= lMin) And (lRet <= lMax) Then
            fValid = True
          Else
            MsgBox "Please enter a number between " & lMin & " and " & lMax, vbCritical
          End If
        Else
          MsgBox "The text you typed is too long", vbCritical
        End If
      Else
        MsgBox "The text you typed is not a valid number", vbCritical
      End If
    End If
  Loop Until (sInput = "") Or fValid
  
  If fValid Then
    LongChooseBox = lRet
  End If
End Function

#If TWINBASIC Then
End Module
#End If
