Attribute VB_Name = "Module2"
Option Explicit
Public Const DLL_PROCESS_DETACH = 0
Public Const DLL_PROCESS_ATTACH = 1
Public Const DLL_THREAD_ATTACH = 2
Public Const DLL_THREAD_DETACH = 3

Public Function DllMain(hInst As Long, fdwReason As Long, lpvReserved As Long) As Boolean
   Select Case fdwReason
      Case DLL_PROCESS_DETACH
         ' No per-process cleanup needed
      Case DLL_PROCESS_ATTACH
         DllMain = True
      Case DLL_THREAD_ATTACH
         ' No per-thread initialization needed
      Case DLL_THREAD_DETACH
         ' No per-thread cleanup needed
   End Select
End Function

Public Function ChkAuthor(var As String) As Boolean
If var <> "Suhas Manjunath" Then
    ChkAuthor = True
Else
    ChkAuthor = False
End If
End Function


