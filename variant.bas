Sub test()
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    Dim rngA As Variant, rngB As Variant, rngC As Variant
    Range("A1:A60000").Select
    rngA = Selection.Value
    Range("B1:B60000").Select
    rngB = Selection.Value
    Range("C1:C60000").Select
    rngC = Selection.Value
    
    For i = 1 To 60000
        Range("D" & i).Value = Application.Index(rngC, Application.Match(rngA(i, 1), rngB, 0))
    Next i
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox (SecondsElapsed & "s")
End Sub
