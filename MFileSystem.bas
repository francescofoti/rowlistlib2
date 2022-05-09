Attribute VB_Name = "MFileSystem"
#If TWINBASIC Then
Module MFileSystem
#End If
#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

'https://francescofoti.com/2017/10/manipulating-filenames-in-excel-or-access-or-vba/
Private Const PATH_SEP      As String = "\"
Private Const PATH_SEP_INV  As String = "/"
Private Const EXT_SEP       As String = "."
Private Const DRIVE_SEP     As String = ":"

Public Function GetFileExt(ByRef psFilename As String) As String
  Dim lLen    As Long
  Dim i       As Long
  Dim sChar   As String

  'Going backwards to find the first EXT_SEP char (or any other path separator)
  lLen = Len(psFilename): i = lLen
  If i Then
    sChar = Mid$(psFilename, i, 1&)
    Do While (i > 0&) And (sChar <> PATH_SEP) And (sChar <> EXT_SEP) And (sChar <> PATH_SEP_INV)
      i = i - 1&: If i = 0& Then Exit Do
      sChar = Mid$(psFilename, i, 1&)
    Loop
    If (i > 0&) And (sChar = EXT_SEP) Then
      GetFileExt = Right$(psFilename, lLen - i)
    End If
  End If
End Function

Public Function StripFileExt(ByRef psFilename As String) As String
  Dim lLen    As Long
  Dim i       As Long
  Dim sChar   As String
  
  lLen = Len(psFilename): i = lLen
  If i Then
    sChar = Mid$(psFilename, i, 1&)
    Do While (i > 0&) And (sChar <> PATH_SEP) And (sChar <> EXT_SEP) And (sChar <> PATH_SEP_INV)
      i = i - 1&: If i = 0& Then Exit Do
      sChar = Mid$(psFilename, i, 1)
    Loop
    If (i > 0) And (sChar = EXT_SEP) Then
      StripFileExt = Left$(psFilename, i - 1&)
    Else
      StripFileExt = psFilename
    End If
  End If
End Function

Public Function StripFilePath(ByVal psFilename As String) As String
  Dim i           As Long
  Dim sChar       As String
  
  i = Len(psFilename)
  If i Then
    sChar = Mid$(psFilename, i, 1)
    While (sChar <> DRIVE_SEP) And (sChar <> PATH_SEP) And (sChar <> PATH_SEP_INV) And (i > 0)
      i = i - 1&
      If i Then
        sChar = Mid$(psFilename, i, 1)
      Else
        sChar = PATH_SEP
      End If
    Wend
    If i Then
      StripFilePath = Right$(psFilename, Len(psFilename) - i)
    Else
      StripFilePath = psFilename
    End If
  End If
End Function

Public Function StripFileName(ByVal psFilename As String) As String
  Dim i           As Long
  Dim fLoop       As Boolean
  Dim sChar       As String * 1
  
  i = Len(psFilename)
  If i Then fLoop = True
  While fLoop
    If i > 0 Then
      sChar = Mid$(psFilename, i, 1)
      If (sChar = PATH_SEP) Or (sChar = DRIVE_SEP) Or (sChar = PATH_SEP_INV) Then fLoop = False
    End If
    If i > 1& Then
      i = i - 1&
    Else
      i = 0&
      fLoop = False
    End If
  Wend
  If i Then
    StripFileName = Left$(psFilename, i)
  Else
    StripFileName = ""
  End If
End Function

Public Function ChangeExt(ByVal psFilename As String, ByVal psNewExt As String) As String
  Dim iLen        As Integer
  Dim i           As Integer
  Dim sChar       As String
  
  If Left$(psNewExt, 1) = EXT_SEP Then psNewExt = Right$(psNewExt, Len(psNewExt) - 1) 'be forgiving
  iLen = Len(psFilename)
  i = iLen
  If i Then
    sChar = Mid$(psFilename, i, 1)
    Do While (i > 0) And (sChar <> PATH_SEP) And (sChar <> EXT_SEP) And (sChar <> PATH_SEP_INV)
      i = i - 1
      If i > 0 Then
        sChar = Mid$(psFilename, i, 1)
      End If
    Loop
    If (i > 0) And (sChar = EXT_SEP) Then
      psFilename = Left$(psFilename, i - 1)
    End If
  End If
  ChangeExt = psFilename & EXT_SEP & psNewExt
End Function

Public Function CombinePath(ByVal psPath1 As String, ByVal psFilename As String) As String
  If (Left$(psFilename, 1) <> PATH_SEP) And (Left$(psFilename, 1) <> PATH_SEP_INV) Then
    CombinePath = NormalizePath(psPath1) & psFilename
  Else
    CombinePath = DenormalizePath(psPath1) & psFilename
  End If
End Function

' Functions NormalizePath, DenormalizePath, ExistFile, ExistDir
' these versions are adapted from  the "Hardcore VB5" book, by Bruce McKinney.

' Make sure path ends in a backslash.
Public Function NormalizePath(ByVal sPath As String) As String
  If (Right$(sPath, 1) <> PATH_SEP) And (Right$(sPath, 1) <> PATH_SEP_INV) Then
    NormalizePath = sPath & PATH_SEP
  Else
    NormalizePath = sPath
  End If
End Function

' Make sure path doesn't end in a backslash
Public Function DenormalizePath(ByVal sPath As String) As String
  If (Right$(sPath, 1) = PATH_SEP) Or (Right$(sPath, 1) = PATH_SEP_INV) Then
    sPath = Left$(sPath, Len(sPath) - 1)
  End If
  DenormalizePath = sPath
End Function

' Test the existence of a file
Public Function ExistFile(psSpec As String) As Boolean
  On Error Resume Next
  Call FileLen(psSpec)
  ExistFile = (Err.Number = 0&)
End Function

Public Function ExistDir(psSpec As String) As Boolean
  On Error Resume Next
  Dim lAttr As Long
  lAttr = GetAttr(psSpec)
  If (Err.Number = 0) Then
    ExistDir = CBool(lAttr And vbDirectory)
  End If
End Function

#If TWINBASIC Then
End Module
#End If

