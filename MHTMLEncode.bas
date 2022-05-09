Attribute VB_Name = "MHTMLEncode"
#If TWINBASIC Then
Module MHTMLEncode
#End If

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

'-- Tableau de correspondance entre code ascii et code html 3.2
Private m_asCode(255)   As String
Private m_asText(255)   As String
Private m_fInitialized  As Boolean

Public Sub InitHTMLEncodeModule()
    'Load arrays
    m_asCode(198) = "AElig"
    m_asCode(193) = "Aacute"
    m_asCode(194) = "Acirc"
    m_asCode(192) = "Agrave"
    m_asCode(197) = "Aring"
    m_asCode(195) = "Atilde"
    m_asCode(196) = "Auml"
    m_asCode(199) = "Ccedil"
    m_asCode(208) = "Dstrok"
    m_asCode(201) = "Eacute"
    m_asCode(202) = "Ecirc"
    m_asCode(200) = "Egrave"
    m_asCode(203) = "Euml"
    m_asCode(205) = "Iacute"
    m_asCode(206) = "Icirc"
    m_asCode(204) = "Igrave"
    m_asCode(207) = "Iuml"
    m_asCode(209) = "Ntilde"
    m_asCode(211) = "Oacute"
    m_asCode(212) = "Ocirc"
    m_asCode(210) = "Ograve"
    m_asCode(216) = "Oslash"
    m_asCode(213) = "Otilde"
    m_asCode(214) = "Ouml"
    m_asCode(222) = "THORN"
    m_asCode(218) = "Uacute"
    m_asCode(219) = "Ucirc"
    m_asCode(217) = "Ugrave"
    m_asCode(220) = "Uuml"
    m_asCode(221) = "Yacute"
    m_asCode(225) = "aacute"
    m_asCode(226) = "acirc"
    m_asCode(180) = "acute"
    m_asCode(230) = "aelig"
    m_asCode(224) = "agrave"
    m_asCode(38) = "amp"
    m_asCode(229) = "aring"
    m_asCode(227) = "atilde"
    m_asCode(228) = "auml"
    m_asCode(166) = "brkbar"
    m_asCode(231) = "ccedil"
    m_asCode(184) = "cedil"
    m_asCode(162) = "cent"
    m_asCode(169) = "copy"
    m_asCode(164) = "curren"
    m_asCode(176) = "deg"
    m_asCode(168) = "die"
    m_asCode(247) = "divide"
    m_asCode(233) = "eacute"
    m_asCode(234) = "ecirc"
    m_asCode(232) = "egrave"
    m_asCode(240) = "eth"
    m_asCode(235) = "euml"
    m_asCode(189) = "fraq12"
    m_asCode(188) = "fraq14"
    m_asCode(190) = "fraq34"
    m_asCode(62) = "gt"
    m_asCode(175) = "hibar"
    m_asCode(237) = "iacute"
    m_asCode(238) = "icirc"
    m_asCode(161) = "iexcl"
    m_asCode(236) = "igrave"
    m_asCode(191) = "iquest"
    m_asCode(239) = "iuml"
    m_asCode(171) = "laqo"
    m_asCode(60) = "lt"
    m_asCode(175) = "macr"
    m_asCode(181) = "micro"
    m_asCode(183) = "middot"
    m_asCode(32) = "nbsp"
    m_asCode(172) = "not"
    m_asCode(241) = "ntilde"
    m_asCode(243) = "oacute"
    m_asCode(244) = "ocirc"
    m_asCode(242) = "ograve"
    m_asCode(170) = "ordf"
    m_asCode(186) = "ordm"
    m_asCode(248) = "oslash"
    m_asCode(245) = "otilde"
    m_asCode(246) = "ouml"
    m_asCode(182) = "para"
    m_asCode(177) = "plusmn"
    m_asCode(163) = "pound"
    m_asCode(34) = "quot"
    m_asCode(187) = "raquo"
    m_asCode(174) = "reg"
    m_asCode(167) = "sect"
    m_asCode(173) = "shy"
    m_asCode(185) = "sup1"
    m_asCode(178) = "sup2"
    m_asCode(179) = "sup3"
    m_asCode(223) = "szlig"
    m_asCode(254) = "thorn"
    m_asCode(215) = "times"
    m_asCode(250) = "uacute"
    m_asCode(251) = "ucirc"
    m_asCode(249) = "ugrave"
    m_asCode(168) = "uml"
    m_asCode(252) = "uuml"
    m_asCode(253) = "yacute"
    m_asCode(165) = "yen"
    m_asCode(255) = "yuml"
    
    m_asText(198) = "AE"
    m_asText(193) = "A"
    m_asText(194) = "A"
    m_asText(192) = "A"
    m_asText(197) = "A"
    m_asText(195) = "A"
    m_asText(196) = "A"
    m_asText(199) = "C"
    m_asText(208) = "D"
    m_asText(201) = "E"
    m_asText(202) = "E"
    m_asText(200) = "E"
    m_asText(203) = "E"
    m_asText(205) = "I"
    m_asText(206) = "I"
    m_asText(204) = "I"
    m_asText(207) = "I"
    m_asText(209) = "N"
    m_asText(211) = "O"
    m_asText(212) = "O"
    m_asText(210) = "O"
    m_asText(216) = "0"
    m_asText(213) = "O"
    m_asText(214) = "O"
    m_asText(222) = ""
    m_asText(218) = "U"
    m_asText(219) = "U"
    m_asText(217) = "U"
    m_asText(220) = "U"
    m_asText(221) = "Y"
    m_asText(225) = "a"
    m_asText(226) = "a"
    m_asText(180) = "a"
    m_asText(230) = "ae"
    m_asText(224) = "a"
    m_asText(229) = "a"
    m_asText(227) = "a"
    m_asText(228) = "a"
    m_asText(231) = "c"
    m_asText(162) = "cent(s)"
    m_asText(169) = "(C)"
    m_asText(176) = "deg"
    m_asText(233) = "e"
    m_asText(234) = "e"
    m_asText(232) = "e"
    m_asText(240) = "eth"
    m_asText(235) = "e"
    m_asText(189) = " 1/2"
    m_asText(188) = " 1/4"
    m_asText(190) = " 3/4"
    m_asText(237) = "i"
    m_asText(238) = "i"
    m_asText(161) = "i"
    m_asText(236) = "i"
    m_asText(191) = "i"
    m_asText(239) = "i"
    m_asText(171) = "<<"
    m_asText(183) = "*"
    m_asText(241) = "n"
    m_asText(243) = "o"
    m_asText(244) = "o"
    m_asText(242) = "o"
    m_asText(170) = "o"
    m_asText(186) = "o"
    m_asText(248) = "0"
    m_asText(245) = "o"
    m_asText(246) = "o"
    m_asText(177) = "+/-"
    m_asText(163) = "pound(s)"
    m_asText(187) = ">>"
    m_asText(174) = "(r)"
    m_asText(185) = "(1)"
    m_asText(178) = "(2)"
    m_asText(179) = "(3)"
    m_asText(223) = "ss"
    m_asText(215) = "*"
    m_asText(250) = "u"
    m_asText(251) = "u"
    m_asText(249) = "u"
    m_asText(168) = "''"
    m_asText(252) = "u"
    m_asText(253) = "y"
    m_asText(165) = "yen(s)"
    m_asText(255) = "y"
    
    m_fInitialized = True
End Sub

Private Function FindCode(s As String) As Integer
  Dim p       As Integer
  Dim iCode   As Integer
  Dim sCode   As String
  Dim i       As Integer
  
  If Not m_fInitialized Then
      InitHTMLEncodeModule
  End If
  
  iCode = -1
  p = InStr(s, ";")
  If p Then
    sCode = Left$(s, p - 1)
    For i = 0 To 255
      If m_asCode(i) = sCode Then
        iCode = i
        Exit For
      End If
    Next i
  End If
    
  FindCode = iCode
End Function

Public Function ToHTML(ByVal s As String) As String
  Dim sRet        As String
  Dim lgSource    As Integer
  Dim i           As Integer
  Dim c           As String
  Dim iCode       As Integer
  
  If Not m_fInitialized Then
      InitHTMLEncodeModule
  End If
  
  On Error Resume Next
  
  lgSource = Len(s)
  If lgSource Then
    For i = 1 To lgSource
      c = Mid$(s, i, 1)
      'cas spécial pour le & qui doit être transformé,
      'sauf s'il désigne un code:
      If c <> "&" Then
        If (c <> " ") And (Asc(c) <> 146) Then '146 c'est un apostrophe "spécial"
          If m_asCode(Asc(c)) <> "" Then
            sRet = sRet & "&" & m_asCode(Asc(c)) & ";"
          Else
            sRet = sRet & c
          End If
        Else
          If c = " " Then
            sRet = sRet & c
          Else
            sRet = sRet & Chr$(39)
          End If
        End If
      Else
        iCode = FindCode(Right$(s, Len(s) - i))
        If iCode <> -1 Then
          sRet = sRet & c
        Else
          sRet = sRet & "&" & m_asCode(Asc(c)) & ";"
        End If
      End If
    Next i
    sRet = Replace(sRet, vbCrLf, "<br>")
    ToHTML = sRet
  End If
End Function

'Replace known html entities by their character
Public Function ToAscii(s As String) As String
  Dim sRet        As String
  Dim lgSource    As Integer
  Dim i           As Integer
  Dim c           As String
  
  If Not m_fInitialized Then
    InitHTMLEncodeModule
  End If
  
  On Error Resume Next
  
  lgSource = Len(s)
  If lgSource Then
    sRet = s
    For i = 0 To 255
      If m_asCode(i) <> "" Then
        sRet = Replace(sRet, "&" & m_asCode(i) & ";", Chr$(i))
      End If
    Next i
    ToAscii = sRet
  End If
End Function

'Arbitrarily translate known characters by their alternate text representation.
'(Note that m_asText is indexed by the Ascii Code, btw Asc not AscW)
Public Function ToText(s As String) As String
  Dim sRet        As String
  Dim lgSource    As Integer
  Dim i           As Integer
  Dim c           As String
  
  If Not m_fInitialized Then
    InitHTMLEncodeModule
  End If
  
  On Error Resume Next
  
  lgSource = Len(s)
  If lgSource Then
    For i = 1 To lgSource
      c = Mid$(s, i, 1)
      If m_asText(Asc(c)) <> "" Then
        sRet = sRet & m_asText(Asc(c))
      Else
        sRet = sRet & c
      End If
    Next i
    ToText = sRet
  End If
End Function

#If TWINBASIC Then
End Module
#End If

