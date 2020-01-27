VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShowMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVersionExA Lib "kernel32" _
            (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Sub ShowMeContextHelp(Cntrl As Control)

   Dim C As Control
   
   Set C = Cntrl
   If C.HelpContextID <> 0 Then
      ShowMeElearning C.HelpContextID
   End If
End Sub

Public Sub ShowMeElearning(Id As Long)

   Dim Url As String
   
   Url = CreateShowMeUrl("ELearning.aspx", Id, 498, 800)
   frmShowMe.ShowNow Url, 13000, 8500
End Sub

Private Function CreateShowMeUrl(PageName As String, HelpCId As Long, Heigt As Integer, Width As Integer)

   Dim Url As String
   
   Url = Client.SysSettings.ShowMeUrl
   If Len(Url) > 0 Then
      If Right(Url, 1) <> "/" Then
         Url = Url & "/"
      End If
      Url = Url & PageName
      
      AddParameterToUrl Url, "CI", CStr(HelpCId)
      AddParameterToUrl Url, "HE", CStr(Heigt)
      AddParameterToUrl Url, "WI", CStr(Width)
      AddParameterToUrl Url, "SN", Client.SysSettings.ShowMeSite
      AddParameterToUrl Url, "UI", CStr(Client.User.UserId)
      AddParameterToUrl Url, "UN", CStr(Client.User.LoginName)
      AddParameterToUrl Url, "CL", CStr(Client.CultureLanguage)
      AddParameterToUrl Url, "AN", App.EXEName
      AddParameterToUrl Url, "AV", ApplicationVersion
      AddParameterToUrl Url, "OV", OSVersion()
      AddParameterToUrl Url, "SI", Client.Station.Id
      
   End If
   CreateShowMeUrl = Url
End Function

Private Sub AddParameterToUrl(ByRef Url, Name As String, Value As String)

   Dim FirstParam As Boolean
   
   If InStr(Url, "?") = 0 Then
      Url = Url & "?"
      FirstParam = True
   End If
   If Not FirstParam Then
      Url = Url & "&"
   End If
   Url = Url & Name & "=" & Value
End Sub
Function OSVersion() As String

   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   On Error Resume Next
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   OSVersion = Format$(osinfo.dwPlatformId, "00") & Format$(osinfo.dwMajorVersion, "00") & Format$(osinfo.dwMinorVersion, "00")
End Function
