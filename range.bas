Sub test()
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    Dim rngA As Range, rngB As Range, rngC As Range
    Set rngA = Range("A1:A60000")
    Set rngB = Range("B1:B60000")
    Set rngC = Range("C1:C60000")
    
    For i = 1 To 60000
        Range("D" & i).Value = Application.Index(rngC, Application.Match(rngA(i), rngB, 0))
    Next i
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox (SecondsElapsed & "s")
End Sub
