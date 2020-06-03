Attribute VB_Name = "modindicator"
Option Explicit

Private IsIndicatorWindowLoaded As Boolean

Public Sub UpdateIndicator()

   Dim NumberOfDictations As Integer

   On Error Resume Next
   NumberOfDictations = NumberOfDictationsForCurrentPatient()
   If NumberOfDictations > 0 Then
      ShowIndicator CStr(NumberOfDictations) & " " & Client.Texts.Txt(1000433, "diktat"), Client.CurrPatient.PatId
   Else
      ShowIndicator "", ""
   End If
End Sub

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

Private Function NumberOfDictationsForCurrentPatient() As Integer

   Dim Dict As clsDict
   Dim TooMany As Boolean
   Dim NumberOfDictations As Integer
   Static LastTimeStamp As Double
   
   LastTimeStamp = Client.DictMgr.CreateList(30005, LastTimeStamp, TooMany)
   Do While Client.DictMgr.ListNextItem(Dict)
      NumberOfDictations = NumberOfDictations + 1
   Loop
   NumberOfDictationsForCurrentPatient = NumberOfDictations
End Function

