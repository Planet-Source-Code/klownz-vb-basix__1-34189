Attribute VB_Name = "Module1"
Public Sub Pause(Duration As Double) 'Used for my loop.
    Dim PotFace As Long
    PotFace = Timer
    Do Until Timer - PotFace >= Duration
    DoEvents
    Loop
End Sub
