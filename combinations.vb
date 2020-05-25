Option Explicit

Sub Sample()
    Dim l As Long, m As Long, n As Long, o As Long, p As Long, q As Long, r As Long, s As Long, t As Long, u As Long
    Dim CountComb As Long, lastrow As Long

    Range("L2").Value = Now

    Application.ScreenUpdating = False

    CountComb = 0: lastrow = 18

For l = 1 To 1
    For m = 1 To 2
        For n = 1 To 2
            For o = 1 To 17
                For p = 1 To 9
                    For q = 1 To 4
                        For r = 1 To 17
                            For s = 1 To 3
                                For t = 1 To 3
                                    For u = 1 To 3
                                        Range("L" & lastrow).Value = Range("A" & l).Value & "/" & _
                                                                     Range("B" & m).Value & "/" & _
                                                                     Range("C" & n).Value & "/" & _
                                                                     Range("D" & o).Value & "/" & _
                                                                     Range("E" & p).Value & "/" & _
                                                                     Range("F" & q).Value & "/" & _
                                                                     Range("G" & r).Value & "/" & _
                                                                     Range("H" & s).Value
                                        lastrow = lastrow + 1
                                        CountComb = CountComb + 1
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
Next

    Range("L1").Value = CountComb
    Range("L3").Value = Now

    Application.ScreenUpdating = True
End Sub

