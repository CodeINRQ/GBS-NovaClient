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

Public Event DeviceChanged()

Public Texts As clsTexts

Public User As clsUser
Public UserMgr As clsUserMgr

Public Station As clsStation
Public StationMgr As clsStationMgr

Public Server As clsServer
Attribute Server.VB_VarHelpID = -1

Public DictTypeMgr As clsDictTypeMgr
Public PriorityMgr As clsPriorityMgr
Public GroupMgr As clsGroupMgr
Public ExtSystemMgr As clsExtSystemMgr
Public EventMgr As clsEventMgr

Public DictMgr As clsDictMgr
Public DictFileMgr As clsDictFileMgr
Public WithEvents PortableMgr As clsPortableMgr
Attribute PortableMgr.VB_VarHelpID = -1

Public OrgMgr As clsOrgMgr
Public RolesMgr As clsRoleMgr

Public StatMgr As clsStatMgr
Public HistMgr As clsHistoryMgr

Public LoggMgr As clsLoggMgr
Public DictAuditMgr As clsDictAuditMgr

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
Public DSSRec As clsDSSRecorder

Public ClientType As ClientTypeEnum
Private mProfileName As String
Private mCTFname As String

Private Sub Class_Initialize()

   SetClientType
   Set Settings = New clsSettings
   SetProfile

   Set Trace = New clsTrace
   Trace.Initialize Settings
   
   Set Texts = New clsTexts
   Texts.CTFPath = CTFPath
   
   SetServer
   Set SysSettings = New clsSysSettings
   Set User = New clsUser
   Set UserMgr = New clsUserMgr
   Set Station = New clsStation
   Set StationMgr = New clsStationMgr
   StationMgr.GetInfo Station
   Set LoggMgr = New clsLoggMgr
   Set DictAuditMgr = New clsDictAuditMgr
   Set DictMgr = New clsDictMgr
   Set DictFileMgr = New clsDictFileMgr
   Set PortableMgr = New clsPortableMgr
   
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
   
   Set DSSRec = New clsDSSRecorder
   Set VoiceXpress = New clsVoiceXpress
End Sub
Public Property Let ProfileName(PName As String)

   mProfileName = PName
   Settings.IniFilePath = ProfileName
End Property
Public Property Get ProfileName() As String

   ProfileName = App.Path & "\" & mProfileName & ".ini"
End Property
Public Property Get ClientName() As String

    Select Case ClientType
       Case ClientType_CareTalk: ClientName = "CareTalk"
       Case ClientType_LegalTalk: ClientName = "LegalTalk"
       Case Else: ClientName = "GrundigNova"
    End Select
End Property
Public Property Get CTFPath() As String

   Dim P As String
   
   Select Case ClientType
      Case ClientType_CareTalk: P = "c"
      Case ClientType_LegalTalk: P = "l"
      Case Else: P = "n"
   End Select
   
   CTFPath = App.Path & "\Txts" & P & ".ctf"
End Property
Private Sub SetServer()

   If Len(Settings.GetString("DictationStorage", "Server", "")) = 0 Then
      Set Server = New clsServerNova
   Else
      Set Server = New clsServerSQL
   End If
   Set Server.Settings = Settings
End Sub
Private Sub SetClientType()

   If UCase(Left(App.EXEName, 4)) = "CARE" Then
      ClientType = ClientType_CareTalk
   ElseIf UCase(Left(App.EXEName, 5)) = "LEGAL" Then
      ClientType = ClientType_LegalTalk
   Else
      ClientType = ClientType_GrundigNova
   End If
End Sub
Private Sub SetProfile()

   ProfileName = CommandValue("Profile", ClientName)
End Sub

Private Sub Class_Terminate()

   Set ExportSettings = Nothing
   Set PortableMgr = Nothing
   Set DSSRec = Nothing
   Set DictMgr = Nothing
   Set StationMgr = Nothing
   Set Station = Nothing
   Set UserMgr = Nothing
   Set User = Nothing
   Set Server = Nothing
   
   Set Trace = Nothing
   Set Settings = Nothing
End Sub

Private Sub PortableMgr_DeviceChanged()

   RaiseEvent DeviceChanged
End Sub

Public Sub UIStatusClear()

   RaiseEvent UIStatusClear
End Sub

Public Sub UIStatusProgress(Total As Long, Left As Long)

   RaiseEvent UIStatusProgress(Total, Left)
End Sub

Public Sub UIStatusSet(StatusText As String, Busy As Boolean)

   RaiseEvent UIStatusSet(StatusText, Busy)
End Sub

Public Sub UIStatusSetSub(SubText As String)

   RaiseEvent UIStatusSetSub(SubText)
End Sub
