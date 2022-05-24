#If Win64 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If
Public Sub custom_wait(finish As Double)
Dim nowtick As Long
Dim endtick As Long
endtick = GetTickCount + (finish * 1000)
Do
nowtick = GetTickCount
DoEvents
Loop Until (nowtick >= endtick) Or (outlooksent = True)

End Sub
