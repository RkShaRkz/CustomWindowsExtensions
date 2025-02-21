'KeepDisplayOn.vb
imports System.Runtime.InteropServices
Public Module MyApplication 
Public Declare UNICODE Function SetThreadExecutionState Lib "Kernel32" (ByVal esFlags as Integer) as Integer
Public Const  ES_AWAYMODE_REQUIRED = &h40
Public Const  ES_CONTINUOUS = &h80000000
Public Const  ES_DISPLAY_REQUIRED = &h2
Public Const  ES_SYSTEM_REQUIRED = &h1
Public Const  ES_USER_PRESENT = &h4

 Public Sub Main ()
  Dim wshshell as Object
  Dim Ret as Integer
  WshShell = CreateObject("WScript.Shell")
  Ret = SetThreadExecutionState(ES_Continuous + ES_Display_Required + ES_Awaymode_Required)
  WshShell.Run(Command(), , True)
 End Sub
End Module