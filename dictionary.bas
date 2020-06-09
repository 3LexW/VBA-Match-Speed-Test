Sub test()
'For early binding, start by opening the Visual Basic Editor by pressing Alt+F11 and going to Tools > References.
'Set a reference to the Microsoft Scripting Runtime object library.
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    Dim rngB As Range, rngC As Range
    Set rngB = Range("B1:B60000")
    Set rngC = Range("C1:C60000")
    
    For i = 1 To 60000
        dict(rngB(i).Value) = rngC(i).Value
    Next i
    
    For i = 1 To 60000
        Range("D" & i).Value = dict(Cells(i, 1).Value)
    Next i
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox (SecondsElapsed & "s")
End Sub
