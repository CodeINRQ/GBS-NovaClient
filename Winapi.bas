Attribute VB_Name = "modWinapi"
Option Explicit

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_ID = (-12)
Private Const WM_GETTEXT = &HD

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" _
                           (ByVal hProcess As Long, ByRef lphModule As Long, _
                            ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                            ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
                           (ByVal hWnd As Long, _
                            ByVal lpClassName As String, _
                            ByVal nMaxCount As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" _
                           (ByVal hProcess As Long, ByVal hModule As Long, _
                            ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, _
                            ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" _
                           (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
                            ByVal dwProcId As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                            ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Function winFindWindowEx(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

   winFindWindowEx = FindWindowEx(hWnd1, hWnd2, lpsz1, lpsz2)
End Function
Public Function winGetClassName(ByVal hWnd As Long) As String

   Const MAXLEN = 255
   Dim ClassName As String
   Dim Ret As Long
   
   ClassName = Space(MAXLEN)
   Ret = GetClassName(hWnd, ClassName, MAXLEN)
   winGetClassName = Left$(ClassName, Ret)
End Function
Public Function winGetForegroundWindow() As Long

   winGetForegroundWindow = GetForegroundWindow()
End Function
Public Function winGetParent(ByVal hWnd As Long) As Long

   winGetParent = GetParent(hWnd)
End Function
'Returns wWnd if no parent window
Public Function winGetTopLevelWindow(ByVal hWnd As Long) As Long

   Dim hParentWindow As Long
   Dim hTemp As Long

   hTemp = hWnd
   Do
      hTemp = GetParent(hTemp)
      If hTemp <> 0 Then
         hParentWindow = hTemp
      End If
   Loop Until hTemp = 0
   
   If hParentWindow <> 0 Then
      winGetTopLevelWindow = hParentWindow
   Else
      winGetTopLevelWindow = hWnd
   End If
End Function
Public Function winGetChildWindowText(ByVal hChildWnd As Long) As String

   Const MAX_LEN = 255
   Dim Caption As String
   Dim Ret As Long
   
   Caption = Space(MAX_LEN)
   Ret = SendMessage(hChildWnd, WM_GETTEXT, MAX_LEN, Caption)
   winGetChildWindowText = Left$(Caption, Ret)

End Function
Public Function winGetWindowControlId(ByVal hWnd As Long) As Long

   winGetWindowControlId = GetWindowLong(hWnd, GWL_ID)
End Function
Public Function winGetWindowModuleName(ByVal hWnd) As String
   
   Const MAX_LEN = 500
   
   Dim Modules(1 To 1) As Long
   Dim Ret As Long
   Dim hProcess As Long
   Dim ModuleName As String
  
   Dim ProcessId As Long

   GetWindowThreadProcessId hWnd, ProcessId
   
   'Get a handle to the Process
   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessId)
   'Got a Process handle
   If hProcess <> 0 Then
      'Get an array of the module handles for the specified process
      Ret = EnumProcessModules(hProcess, Modules(1), 1, 0)
      'If the Module Array is retrieved, Get the ModuleFileName
      If Ret <> 0 Then
         ModuleName = Space(MAX_LEN)
         Ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, MAX_LEN)
         ModuleName = Left$(ModuleName, Ret)
      End If
      'Close the handle to the process
      Ret = CloseHandle(hProcess)
   End If
   winGetWindowModuleName = ModuleName
End Function
Public Function winGetWindowRect(ByVal hWnd As Long, Rect As Rect) As Long

  winGetWindowRect = GetWindowRect(hWnd, Rect)
End Function
Public Function winGetClientRect(ByVal hWnd As Long, Rect As Rect) As Long

  winGetClientRect = GetClientRect(hWnd, Rect)
End Function
Public Function winGetWindowText(ByVal hWnd As Long) As String

   Const MAXLEN = 255
   Dim Caption As String
   Dim Ret As Long
   
   Caption = Space(MAXLEN)
   Ret = GetWindowText(hWnd, Caption, MAXLEN)
   If Ret > 0 Then
      winGetWindowText = Left$(Caption, Ret)
   End If
End Function
