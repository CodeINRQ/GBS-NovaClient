Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function HotKey(ByVal KeyCode As Integer) As Boolean

    'Call it once to clear its old data
    'GetAsyncKeyState KeyCode

    'Now actually check for the key
    If GetAsyncKeyState(KeyCode) Then
        HotKey = True
    Else
        HotKey = False
    End If

End Function
