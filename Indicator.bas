Attribute VB_Name = "modindicator"
Option Explicit

Private IsIndicatorWindowLoaded As Boolean

Public Sub ShowIndicator(Tip As String, Id As String)

   If Client.SysSettings.IndicatorActive Then
      If Len(Tip) > 0 Then
         LoadIndicator
         frmIndicator.SetIndicatorText Tip, Id
      Else
         If IsIndicatorWindowLoaded Then
            frmIndicator.SetIndicatorText "", ""
         End If
      End If
   End If
End Sub
Public Sub UnloadIndicator()

   If IsIndicatorWindowLoaded Then
      WindowNotFloating frmIndicator
   End If
   Unload frmIndicator
End Sub
Public Sub LoadIndicator()

   If Not IsIndicatorWindowLoaded Then
      IsIndicatorWindowLoaded = True
      frmIndicator.Move Screen.Width - frmIndicator.Width - 1200, 60
      WindowFloating frmIndicator
   End If
End Sub
