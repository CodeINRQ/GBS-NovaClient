Attribute VB_Name = "modFönster"
Option Explicit

Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetTickCount Lib "kernel32" () As Long

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Sub FlytandeFönster(F As Form)

   frmMain.imgPin.Visible = False
   frmMain.imgPinin.Visible = True
   Call SetWindowPos(F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
Sub IckeFlytandeFönster(F As Form)

   frmMain.imgPin.Visible = True
   frmMain.imgPinin.Visible = False
   Call SetWindowPos(F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Sub LäsFönsterläge(F As Form)

   Dim S As String
   Dim P As Integer
   Dim WinState As Integer

   On Error Resume Next
   S = GetIniString("Fönster", "Läge", "")
   If S <> "" Then
      P = InStr(S, ",")
      WinState = Heltal(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Top = Heltal(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Left = Heltal(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Height = Heltal(Left(S, P - 1))
      S = Mid$(S, P + 1)
      P = InStr(S, ",")
      F.Width = Heltal(S)
      F.WindowState = WinState
   Else
      F.Left = (Screen.Width - F.Width) / 2
      F.Top = (Screen.Height - F.Height) / 2
   End If
End Sub

Sub SparaFönsterläge(F As Form)

   Dim WinState As Integer
   Dim S As String

   On Error Resume Next
   WinState = F.WindowState
   F.WindowState = 0
   S = WinState & "," & F.Top & "," & F.Left & "," & F.Height & "," & F.Width
   F.Hide
   WriteIniString "Fönster", "Läge", S
End Sub

