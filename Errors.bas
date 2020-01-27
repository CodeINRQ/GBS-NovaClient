Attribute VB_Name = "modErrors"
Option Explicit

Public Sub ErrorHandle(Location As String, Eno As Long, UserMsgNr As Long, UserMsgDef As String, EndProgram As Boolean)

   On Error Resume Next
   Client.Trace.AddRow Trace_Level_FatalErrors, "EH", Location, "EH", Eno, UserMsgNr
   Client.LoggMgr.Insert 1320114, LoggLevel_SysFailure, 0, CStr(UserMsgNr) & " (" & Location & ":" & CStr(Eno) & ")"
   MsgBox Client.Texts.Txt(UserMsgNr, UserMsgDef) & " (" & Location & ":" & CStr(Eno) & ")"
   If EndProgram Then
      Unload frmMain
      End
   End If
End Sub
Public Sub ErrorHandleExplicit(Location As String, Desc As String, UserMsgNr As Long, UserMsgDef As String, EndProgram As Boolean)

   On Error Resume Next
   Client.LoggMgr.Insert 1320115, LoggLevel_SysFailure, 0, CStr(UserMsgNr) & " (" & Location & ":0)"
   MsgBox Client.Texts.Txt(UserMsgNr, UserMsgDef) & " (" & Location & ":" & CStr(Desc) & ")"
   If EndProgram Then
      Unload frmMain
      End
   End If
End Sub

