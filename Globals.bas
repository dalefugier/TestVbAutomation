Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function CoAllowSetForegroundWindow Lib "ole32.dll" (ByVal pUnk As Object, ByVal lpvReserved As Long) As Long


