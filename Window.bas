Attribute VB_Name = "modWindow"
Option Explicit


Declare Function SetWindowPos Lib "user32" _
   (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private LastForegroundWindow As Long

Sub SetWindowTopMostAndForeground(F As Form)
   
   If F.WindowState = vbMinimized Then
      F.WindowState = vbNormal
   End If
   WindowFloating F
   WindowNotFloating F

   SetNewForgroundWindow F.hWnd

End Sub
Function SaveForegroundWindow() As Long

   LastForegroundWindow = winGetForegroundWindow()
   
   SaveForegroundWindow = LastForegroundWindow
End Function
Public Function FindControlOnWindow(hWnd As Long, ControlId As Long, ByRef Caption As String, ByRef hControl As Long) As Boolean

   Dim Ret As Long
   Dim Rct As Rect
   
   Do
      hControl = winFindWindowEx(hWnd, hControl, vbNullString, vbNullString)
      If hControl <> 0 Then
         If winGetWindowControlId(hControl) = ControlId Then
            FindControlOnWindow = True
            
            Caption = winGetChildWindowText(hControl)
               
            Exit Function
         End If
      End If
   Loop Until hControl = 0
   
End Function
Public Function GetWindowRectAsString(hWnd As Long) As String

   Dim Rectangle As Rect
   Dim Ret As Long
   
   Ret = winGetWindowRect(hWnd, Rectangle)
   If Ret <> 0 Then
      GetWindowRectAsString = FormatRectAsString(Rectangle)
   End If
End Function
Public Function GetClientRectAsString(hWnd As Long) As String

   Dim Rectangle As Rect
   Dim Ret As Long
   
   Ret = winGetClientRect(hWnd, Rectangle)
   If Ret <> 0 Then
      GetClientRectAsString = FormatRectAsString(Rectangle)
   End If
End Function

Public Function FormatRectAsString(Rect As Rect) As String

   FormatRectAsString = CStr(Rect.Left) & "," & CStr(Rect.Top) & "," & CStr(Rect.Right) & "," & CStr(Rect.Bottom)
End Function
Function RestoreForegroundWindow(Optional WindowHandle As Long = 0) As Long
  
   If WindowHandle <> 0 Then
      LastForegroundWindow = WindowHandle
   End If
   If LastForegroundWindow <> 0 Then
'      Client.Trace.AddRow Trace_Level_Full, "Win", "RFW", "FGWC", CStr(LastForegroundWindow)
      SetNewForgroundWindow LastForegroundWindow
      LastForegroundWindow = 0
   End If
End Function
Private Sub SetNewForgroundWindow(WindowHandle As Long)

   'Dim WCaption As String
   Dim Res As Boolean

   'WCaption = GetForegroundWindowCaption(WindowHandle)
   'Res = SetForegroundWindow(WindowHandle)
   Res = ForceForegroundWindow(WindowHandle)
'   Client.Trace.AddRow Trace_Level_Full, "Win", "SNFW", "SNFW", WCaption, CStr(Res)
End Sub
Sub WindowFloating(F As Form, Optional SetAlsoForeground As Boolean = False)

   Call SetWindowPos(F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
   If SetAlsoForeground Then
      SetNewForgroundWindow F.hWnd
   End If
End Sub
Sub WindowNotFloating(F As Form)

   Call SetWindowPos(F.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Sub WindowSetPositionFromString(F As Form, S)

   Dim P As Integer
   Dim WinState As Integer

   On Error Resume Next
   If S <> "" Then
      P = InStr(S, ",")
      WinState = CLng(Left(S, P - 1))
      S = mId$(S, P + 1)
      P = InStr(S, ",")
      F.Top = CLng(Left(S, P - 1))
      S = mId$(S, P + 1)
      P = InStr(S, ",")
      F.Left = CLng(Left(S, P - 1))
      S = mId$(S, P + 1)
      P = InStr(S, ",")
      F.Height = CLng(Left(S, P - 1))
      S = mId$(S, P + 1)
      P = InStr(S, ",")
      F.Width = CLng(S)
      F.WindowState = WinState
   Else
      F.Left = (Screen.Width - F.Width) / 2
      F.Top = (Screen.Height - F.Height) / 2
   End If
End Sub

Function WindowSavePositionToString(F As Form) As String

   Dim WinState As Integer

   On Error Resume Next
   WinState = F.WindowState
   F.WindowState = 0
   WindowSavePositionToString = WinState & "," & F.Top & "," & F.Left & "," & F.Height & "," & F.Width
   F.WindowState = WinState
End Function

Sub CenterForm(F As Form, BaseFomr As Form)

   Dim I As Integer

   If BaseFomr.WindowState <> vbMinimized Then
      I = (BaseFomr.Left + BaseFomr.Width / 2) - F.Width / 2
      If I < 0 Then I = 0
      If I > Screen.Width - F.Width Then I = Screen.Width - F.Width
      F.Left = I
      I = (BaseFomr.Top + BaseFomr.Height / 2) - F.Height / 2
      If I < 0 Then I = 0
      If I > Screen.Height - F.Height Then I = Screen.Height - F.Height
      F.Top = I
   End If
End Sub
Sub CenterAndTranslateForm(F As Form, BaseFomr As Form)

   CenterForm F, BaseFomr
   Client.Texts.ApplyToOneForm F
End Sub
Sub TranslateForm(F As Form)

   Client.Texts.ApplyToOneForm F
End Sub
Public Function ForceForegroundWindow(ByVal hWnd As Long) As Boolean

   Dim ThreadID1 As Long
   Dim ThreadID2 As Long
   Dim nRet As Long
   
   ' Nothing to do if already in foreground.
   If hWnd = winGetForegroundWindow() Then
      ForceForegroundWindow = True
   Else
      ' First need to get the thread responsible for
      ' the foreground window, then the thread running
      ' the passed window.
      ThreadID1 = GetWindowThreadProcessId(winGetForegroundWindow, ByVal 0&)
      ThreadID2 = GetWindowThreadProcessId(hWnd, ByVal 0&)
      
      ' By sharing input state, threads share their
      ' concept of the active window.
      If ThreadID1 <> ThreadID2 Then
         Call AttachThreadInput(ThreadID1, ThreadID2, True)
         nRet = SetForegroundWindow(hWnd)
         Call AttachThreadInput(ThreadID1, ThreadID2, False)
      Else
         nRet = SetForegroundWindow(hWnd)
      End If
      
      ' Restore and repaint
      'If IsIconic(hWnd) Then
      '   Call ShowWindow(hWnd, SW_RESTORE)
      'Else
      '   Call ShowWindow(hWnd, SW_SHOW)
      'End If
      
      ' SetForegroundWindow return accurately reflects success.
      ForceForegroundWindow = CBool(nRet)
   End If
End Function



