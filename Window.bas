Attribute VB_Name = "modWindow"
Option Explicit

Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Sub WindowFloating(F As Form)

   Call SetWindowPos(F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
Sub WindowNotFloating(F As Form)

   Call SetWindowPos(F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Sub WindowSetPositionFromString(F As Form, S)

   Dim P As Integer
   Dim WinState As Integer

   On Error Resume Next
   If S <> "" Then
      P = InStr(S, ",")
      WinState = CLng(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Top = CLng(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Left = CLng(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Height = CLng(Left(S, P - 1))
      S = Mid$(S, P + 1)
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

