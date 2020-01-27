Attribute VB_Name = "modVoiceExpress"
Option Explicit

Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) _
   As Long

Private Declare Function SendInput Lib "user32.dll" ( _
    ByVal nInputs As Long, _
    pInputs As GENERALINPUT, _
    ByVal cbSize As Long _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal ByteLen As Long _
)

Private Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type

Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Const INPUT_KEYBOARD = 1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5

Private Sub SendKey(bKey As Byte, UpDown As Integer)
    Dim GInput As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  ' the key we're going to release
    KInput.dwFlags = UpDown  ' release the key
    GInput.dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput.xi(0), KInput, Len(KInput)
    Call SendInput(1, GInput, Len(GInput))
End Sub
Public Sub VoiceExpress(Start As Boolean)

   'SendKey VK_LMENU, 0
   'SendKey VK_LSHIFT, 0
   'delay 0.3
   'SendKey VK_LSHIFT, KEYEVENTF_KEYUP
   'SendKey VK_LMENU, KEYEVENTF_KEYUP
   
   
End Sub
Private Sub delay(Sec As Double)

   Dim T As Double
   
   T = Timer + Sec
   Do While T > Timer
      DoEvents
   Loop
End Sub
Public Function IsVoiceExpressRunning() As Boolean

   Dim Res As Long
   
   Res = FindWindow(vbNullString, "Voice Xpress")
End Function

Public Sub test()
   Dim x As New SPEECHCENTERLib.SCApplication
   
   x.AppBarVisible = 0
   x.Listening = 2
End Sub
