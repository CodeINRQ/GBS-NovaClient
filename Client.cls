VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event UIStatusSet(StatusText As String, Busy As Boolean)
Public Event UIStatusSetSub(SubText As String)
Public Event UIStatusProgress(Total As Long, Left As Long)
Public Event UIStatusClear()

Public Texts As clsTexts

Public User As clsUser
Public UserMgr As clsUserMgr

Public Station As clsStation
Public StationMgr As clsStationMgr

Public WithEvents Server As clsServerSQL
Attribute Server.VB_VarHelpID = -1
'Public WithEvents Server As clsServerJet

Public DictTypeMgr As clsDictTypeMgr
Public PriorityMgr As clsPriorityMgr
Public GroupMgr As clsGroupMgr
Public ExtSystemMgr As clsExtSystemMgr
Public EventMgr As clsEventMgr

Public DictMgr As clsDictMgr

Public OrgMgr As clsOrgMgr
Public RolesMgr As clsRoleMgr

Public StatMgr As clsStatMgr
Public HistMgr As clsHistoryMgr

Public LoggMgr As clsLoggMgr
Public DictAuditMgr As clsDictAuditMgr

Public DSSRec As CareTalkDSSRec3.DSSRecorder
Public DSSRecConnector As CareTalkDSSRec3.Connector
Public Hw As Gru_Harware

Public VoiceXpress As clsVoiceXpress

Public Settings As clsSettings
Public Trace As clsTrace

Public SysSettings As New clsSysSettings
Public CultureLanguage As String
Public NewRecInfo As New clsNewRecInfo
Public CurrPatient As New clsCurrPat
Public PlayDictIdFromAPI As Long

Public ExportSettings As clsExportSettings

Public WindowHandleForIndicator As Long

Private Sub Class_Initialize()

   Set Settings = New clsSettings
   Settings.IniFilePath = App.Path & "\" & CommandValue("Profile", "CareTalk") & ".ini"

   Set Trace = New clsTrace
   Trace.Initialize Settings
   
   Set Texts = New clsTexts
   'Set Server = New clsServerJet
   Set Server = New clsServerSQL
   Set SysSettings = New clsSysSettings
   Set User = New clsUser
   Set UserMgr = New clsUserMgr
   Set Station = New clsStation
   Set StationMgr = New clsStationMgr
   StationMgr.GetInfo Station
   Set LoggMgr = New clsLoggMgr
   Set DictAuditMgr = New clsDictAuditMgr
   Set DictMgr = New clsDictMgr
   Set GroupMgr = New clsGroupMgr
   Set DictTypeMgr = New clsDictTypeMgr
   Set PriorityMgr = New clsPriorityMgr
   Set ExtSystemMgr = New clsExtSystemMgr
   Set EventMgr = New clsEventMgr
   Set OrgMgr = New clsOrgMgr
   Set RolesMgr = New clsRoleMgr
   Set StatMgr = New clsStatMgr
   Set HistMgr = New clsHistoryMgr
   Set ExportSettings = New clsExportSettings
   
   Set DSSRecConnector = New CareTalkDSSRec3.Connector
   Set DSSRec = DSSRecConnector.DSSRecorder
   DSSRec.Play
   DSSRec.Rec
   'DSSRec.Initialize ""
   
   Set VoiceXpress = New clsVoiceXpress
End Sub

Private Sub Class_Terminate()

   Set ExportSettings = Nothing
   Set DSSRec = Nothing
   Set DSSRecConnector = Nothing
   Set DictMgr = Nothing
   Set StationMgr = Nothing
   Set Station = Nothing
   Set UserMgr = Nothing
   Set User = Nothing
   Set Server = Nothing
   
   Set Trace = Nothing
   Set Settings = Nothing
End Sub

Private Sub Server_UIStatusClear()

   RaiseEvent UIStatusClear
End Sub

Private Sub Server_UIStatusProgress(Total As Long, Left As Long)

   RaiseEvent UIStatusProgress(Total, Left)
End Sub

Private Sub Server_UIStatusSet(StatusText As String, Busy As Boolean)

   RaiseEvent UIStatusSet(StatusText, Busy)
End Sub

Private Sub Server_UIStatusSetSub(SubText As String)

   RaiseEvent UIStatusSetSub(SubText)
End Sub
