Private Function NormaliseKey(ByVal String1 As String) As String
NormaliseKey = Replace(UCase$(String1), " ", "")
End Function



Function FuzzyCount(ByVal LookupValue As String, _
                      ByVal TableArray As Range, _
                      Optional NFPercent As Single = 0.05, _
                      Optional Algorithm As Variant = 3) As Long
'**********************************************************************
'** Simple count of (Fuzzy) Matching strings >= NFPercent threshold  **
'**********************************************************************
Dim lMatchCount As Long

Dim rCur As Range

Dim sString1 As String
Dim sString2 As String

'** Normalise lookup value **
sString1 = LCase$(Application.Trim(LookupValue))

For Each rCur In Intersect(TableArray.Resize(, 1), Sheets(TableArray.Parent.Name).UsedRange)

    '** Normalise current Table entry **
    sString2 = LCase$(Application.Trim(CStr(rCur)))

    If sString2 <> "" Then
        If FuzzyPercent(String1:=sString1, _
                        String2:=sString2, _
                        Algorithm:=Algorithm, _
                        Normalised:=False) >= NFPercent Then
            lMatchCount = lMatchCount + 1
        End If
    End If
Next rCur

FuzzyCount = lMatchCount

End Function

Function FuzzyPercent(ByVal String1 As String, _
                      ByVal String2 As String, _
                      Optional Algorithm As Variant = 3, _
                      Optional Normalised As Boolean = False) As Single
'*************************************
'** Return a % match on two strings **
'*************************************
Dim bSoundex As Boolean
Dim bBasicMetaphone As Boolean
Dim intLen1 As Integer, intLen2 As Integer
Dim intCurLen As Integer
Dim intTo As Integer
Dim intPos As Integer
Dim intPtr As Integer
Dim intScore As Integer
Dim intTotScore As Integer
Dim intStartPos As Integer
Dim lngAlgorithm As Long
Dim sngScore As Single
Dim strWork As String

bSoundex = LCase$(CStr(Algorithm)) = "soundex"
bBasicMetaphone = LCase$(CStr(Algorithm)) = "metaphone"

'-------------------------------------------------------
'-- If strings havent been normalised, normalise them --
'-------------------------------------------------------
If Normalised = False Then
    If bSoundex Or bBasicMetaphone Then
       String1 = NormaliseStringAtoZ(String1)
       String2 = NormaliseStringAtoZ(String2)
    Else
        String1 = LCase$(Application.Trim(String1))
        String2 = LCase$(Application.Trim(String2))
    End If
End If

'----------------------------------------------
'-- Give 100% match if strings exactly equal --
'----------------------------------------------
If String1 = String2 Then
    FuzzyPercent = 1
    Exit Function
End If

'If bSoundex Then
'    String1 = Soundex(Replace(String1, " ", ""))
'    String2 = Soundex(Replace(String2, " ", ""))
'    If String1 = String2 Then
'        FuzzyPercent = msngSoundexMatchPercent
'    Else
'        FuzzyPercent = 0
'    End If
'    Exit Function
'ElseIf bBasicMetaphone Then
'    String1 = Metaphone1(String1)
'    String2 = Metaphone1(String2)
'    If String1 = String2 Then
'        FuzzyPercent = msngMetaphoneMatchPercent
'    Else
'        FuzzyPercent = 0
'    End If
'    Exit Function
'End If

intLen1 = Len(String1)
intLen2 = Len(String2)

If intLen1 = 0 Or intLen2 = 0 Then
    FuzzyPercent = 0
    Exit Function
End If

'----------------------------------------
'-- Give 0% match if string length < 2 --
'----------------------------------------
If intLen1 < 2 Then
    FuzzyPercent = 0
    Exit Function
End If

intTotScore = 0                   'initialise total possible score
intScore = 0                      'initialise current score

lngAlgorithm = Val(Algorithm)

'--------------------------------------------------------
'-- If Algorithm = 1 or 3, Search for single characters --
'--------------------------------------------------------
If (lngAlgorithm And 1) <> 0 Then
    If intLen1 < intLen2 Then
        FuzzyAlg1 String1, String2, intScore, intTotScore
    Else
        FuzzyAlg1 String2, String1, intScore, intTotScore
    End If
End If

'-----------------------------------------------------------
'-- If Algorithm = 2 or 3, Search for pairs, triplets etc. --
'-----------------------------------------------------------
If (lngAlgorithm And 2) <> 0 Then
    If intLen1 < intLen2 Then
        FuzzyAlg2 String1, String2, intScore, intTotScore
    Else
        FuzzyAlg2 String2, String1, intScore, intTotScore
    End If
End If

'-------------------------------------------------------------
'-- If Algorithm = 4,5,6,7, use Levenstein Distance method  --
'-- (Algorithm 4 was Dan Ostrander's code)                  --
'-------------------------------------------------------------
If (lngAlgorithm And 4) <> 0 Then
    If intLen1 < intLen2 Then
'        sngScore = FuzzyAlg4(String1, String1)
        sngScore = GetLevenshteinPercentMatch(String1:=String1, _
                                              String2:=String2, _
                                              Normalised:=True)
    Else
'        sngScore = FuzzyAlg4(String2, String1)
        sngScore = GetLevenshteinPercentMatch(String1:=String2, _
                                              String2:=String1, _
                                              Normalised:=True)
    End If
    intScore = intScore + (sngScore * 100)
    intTotScore = intTotScore + 100
End If

FuzzyPercent = intScore / intTotScore

End Function

Private Sub FuzzyAlg1(ByVal String1 As String, _
                      ByVal String2 As String, _
                      ByRef Score As Integer, _
                      ByRef TotScore As Integer)
Dim intLen1 As Integer, intPos As Integer, intPtr As Integer, intStartPos As Integer

intLen1 = Len(String1)
TotScore = TotScore + intLen1              'update total possible score
intPos = 0
For intPtr = 1 To intLen1
    intStartPos = intPos + 1
    intPos = InStr(intStartPos, String2, Mid$(String1, intPtr, 1))
    If intPos > 0 Then
        If intPos > intStartPos + 3 Then     'No match if char is > 3 bytes away
            intPos = intStartPos
        Else
            Score = Score + 1          'Update current score
        End If
    Else
        intPos = intStartPos
    End If
Next intPtr
End Sub
Private Sub FuzzyAlg2(ByVal String1 As String, _
                        ByVal String2 As String, _
                        ByRef Score As Integer, _
                        ByRef TotScore As Integer)
Dim intCurLen As Integer, intLen1 As Integer, intTo As Integer, intPtr As Integer, intPos As Integer
Dim strWork As String

intLen1 = Len(String1)
For intCurLen = 1 To intLen1
    strWork = String2                          'Get a copy of String2
    intTo = intLen1 - intCurLen + 1
    TotScore = TotScore + Int(intLen1 / intCurLen)  'Update total possible score
    For intPtr = 1 To intTo Step intCurLen
        intPos = InStr(strWork, Mid$(String1, intPtr, intCurLen))
        If intPos > 0 Then
            Mid$(strWork, intPos, intCurLen) = String$(intCurLen, &H0) 'corrupt found string
            Score = Score + 1     'Update current score
        End If
    Next intPtr
Next intCurLen

End Sub
'Private Function FuzzyAlg4(strIn1 As String, strIn2 As String) As Single
'
'Dim L1               As Integer
'Dim In1Mask(1 To 24) As Long     'strIn1 is 24 characters max
'Dim iCh              As Integer
'Dim N                As Long
'Dim strTry           As String
'Dim strTest          As String
'
'TopMatch = 0
'L1 = Len(strIn1)
'strTest = UCase(strIn1)
'strCompare = UCase(strIn2)
'For iCh = 1 To L1
'    In1Mask(iCh) = 2 ^ iCh
'Next iCh      'Loop thru all ordered combinations of characters in strIn1
'For N = 2 ^ (L1 + 1) - 1 To 1 Step -1
'    strTry = ""
'    For iCh = 1 To L1
'        If In1Mask(iCh) And N Then
'            strTry = strTry & Mid(strTest, iCh, 1)
'        End If
'    Next iCh
'    If Len(strTry) > TopMatch Then FuzzyAlg4Test strTry
'Next N
'FuzzyAlg4 = TopMatch / CSng(L1)
'End Function
'Sub FuzzyAlg4Test(strIn As String)
'
'Dim l          As Integer
'Dim strTry   As String
'Dim iCh        As Integer
'
'l = Len(strIn)
'If l <= TopMatch Then Exit Sub
'strTry = "*"
'For iCh = 1 To l
'    strTry = strTry & Mid(strIn, iCh, 1) & "*"
'Next iCh
'If strCompare Like strTry Then
'    If l > TopMatch Then TopMatch = l
'End If
'End Sub

Public Function GetLevenshteinPercentMatch(ByVal String1 As String, _
                                            ByVal String2 As String, _
                                            Optional Normalised As Boolean = False) As Single
Dim iLen As Integer
If Normalised = False Then
    String1 = UCase$(WorksheetFunction.Trim(String1))
    String2 = UCase$(WorksheetFunction.Trim(String2))
End If
iLen = WorksheetFunction.Max(Len(String1), Len(String2))
GetLevenshteinPercentMatch = (iLen - LevenshteinDistance(String1, String2)) / iLen
End Function

Private Function NormaliseStringAtoZ(ByVal String1 As String) As String
'---------------------------------------------------------
'-- Remove all but alpha chars and convert to lowercase --
'---------------------------------------------------------
Dim iPtr As Integer
Dim sChar As String
Dim sResult As String

sResult = ""
For iPtr = 1 To Len(String1)
    sChar = LCase$(Mid$(String1, iPtr, 1))
    If sChar <> UCase$(sChar) Then sResult = sResult & sChar
Next iPtr
NormaliseStringAtoZ = sResult
End Function

'********************************
'*** Compute Levenshtein Distance
'********************************

Public Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Integer
Dim d() As Integer ' matrix
Dim m As Integer ' length of t
Dim N As Integer ' length of s
Dim I As Integer ' iterates through s
Dim j As Integer ' iterates through t
Dim s_i As String ' ith character of s
Dim t_j As String ' jth character of t
Dim cost As Integer ' cost

  ' Step 1

  N = Len(s)
  m = Len(t)
  If N = 0 Then
    LevenshteinDistance = m
    Exit Function
  End If
  If m = 0 Then
    LevenshteinDistance = N
    Exit Function
  End If
  ReDim d(0 To N, 0 To m) As Integer

  ' Step 2

  For I = 0 To N
    d(I, 0) = I
  Next I

  For j = 0 To m
    d(0, j) = j
  Next j

  ' Step 3

  For I = 1 To N

    s_i = Mid$(s, I, 1)

    ' Step 4

    For j = 1 To m

      t_j = Mid$(t, j, 1)

      ' Step 5

      If s_i = t_j Then
        cost = 0
      Else
        cost = 1
      End If

      ' Step 6

      d(I, j) = WorksheetFunction.Min(d(I - 1, j) + 1, d(I, j - 1) + 1, d(I - 1, j - 1) + cost)

    Next j

  Next I

  ' Step 7

  LevenshteinDistance = d(N, m)

End Function