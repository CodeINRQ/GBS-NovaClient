VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Event UIStatusSet(StatusText As String, Busy As Boolean)
'Public Event UIStatusSetSub(SubText As String)
'Public Event UIStatusProgress(Total As Long, Left As Long)
'Public Event UIStatusClear()

Public Property Get SQLDateStringFormat() As String
End Property
Public Property Get SQLDateDelimiter() As String
End Property
Public Property Get SQLDateTimeStringFormat() As String
End Property
Public Property Get Server() As String
End Property
Public Property Get Database() As String
End Property
Public Property Get StorageOpened() As Boolean
End Property
Public Property Set Settings(s As clsSettings)
End Property
Public Function IsBatchDoneByClient() As Boolean
End Function
Public Sub DeleteAllDictations()
End Sub
Public Sub DeleteHistory(DeleteOlderThanDays As Integer)
End Sub
Public Function GetCurrentTimestamp() As Date
End Function
Private Sub AddHistory(ByRef Dict As clsDict)
End Sub
Private Sub DeleteHistoryByDictId(DictId As Long)
End Sub
Public Function GetRoles(Roles As clsRoles, ByVal GroupId As Long, ByVal OrgId As Long) As Boolean
End Function
Public Function GetSDKLic(SDKLic As clsSDKLicense, ByVal Site As String, ByVal StationId As String) As Boolean
End Function
Public Sub SaveRoles(Roles As clsRoles)
End Sub
Public Sub SaveOrg(Org As clsOrg)
End Sub
Public Sub SaveDictType(DictType As clsDictType)
End Sub
Public Sub SaveGroup(Grp As clsGroup)
End Sub
Public Sub SaveUser(Usr As clsUser)
End Sub
Public Sub SaveUserHomeOrg(Usr As clsUser)
End Sub
Public Sub ActivateUser(Usr As clsUser, Deactivate As Boolean)
End Sub
Public Sub DeleteOneUserGroup(UserId As Long, GroupId As Long)
End Sub
Public Sub DeleteAllUserGroup(UserId As Long)
End Sub
Public Sub SaveUserGroup(UserId As Long, GroupId As Long)
End Sub
Public Sub CreateDictIdList(OrgId As Long, AuthorId As Long, TranscriberId As Long, Optional PatId As String = "", Optional PatId2 As String = "")
End Sub
Public Function DictIdGetNext(ByRef DictId As Long) As Boolean
End Function
Public Sub CreateHist(OrgId As Long, HistYear As Integer, HistType As HistTypeEnum)
End Sub
Public Function HistGetNext(ByRef Hist As clsHistory) As Boolean
End Function
Private Sub ConvertHistRsToClass(ByRef Rs As ADODB.Recordset, ByRef Hist As clsHistory)
End Sub
Public Sub CreateStat(OrgId As Long, ByRef Stat As clsStat, StatusIdStart As Long, StatusIdEnd As Long, _
                      DaysLimit1 As Integer, _
                      DaysLimit2 As Integer, _
                      DaysLimit3 As Integer, _
                      DaysLimit4 As Integer, _
                      DaysLimit5 As Integer, _
                      DaysLimit6 As Integer, _
                      DaysLimit7 As Integer)
End Sub
Public Function GetSoundfilePath(DictId As Long) As String
End Function
Public Function NeedLoginInfo() As Boolean
End Function
Public Function UserLogin(User As clsUser, _
                          LoginName As String, _
                          Password As String, _
                          NewPassword As String, _
                          ExtSystem As String, _
                          ExtPassword As String) As Integer
End Function
Private Function SetUserFromSQLLoginName(User As clsUser) As Boolean
End Function
Private Function ChangeSQLPassword(User As clsUser, OldPassword As String, NewPassword As String)
End Function
Public Sub CreateLoggList(StartTime As Date, EndTime As Date, StartLevel As Integer, EndLevel As Integer)
End Sub
Public Function LoggListGetNext(ByRef Logg As clsLogg) As Boolean
End Function
Private Sub ConvertLoggRsToClass(ByRef Rs As ADODB.Recordset, ByRef Logg As clsLogg)
End Sub
Public Sub LoggInsert(Logg As clsLogg)
End Sub
Public Sub CreateDictAuditList(DictId As Long)
End Sub
Public Function DictAuditGetNext(ByRef Audit As clsDictAudit) As Boolean
End Function
Private Sub ConvertDictAuditRsToClass(ByRef Rs As ADODB.Recordset, ByRef Audit As clsDictAudit)
End Sub
Public Sub CreateDictAuditForUserList(UserId As Long)
End Sub
Public Function DictAuditForUserGetNext(ByRef Audit As clsDictAudit) As Boolean
End Function
Private Sub ConvertDictAuditForUserRsToClass(ByRef Rs As ADODB.Recordset, ByRef Audit As clsDictAudit)
End Sub
Public Sub CreateExtSystemList()
End Sub
Public Function ExtSystemGetNext(ByRef E As clsExtSystem) As Boolean
End Function
Private Sub ConvertExtSystemRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystem)
End Sub
Public Sub CreateExtSystemDictTypeList()
End Sub
Public Function ExtSystemDictTypeGetNext(ByRef E As clsExtSystemDictType) As Boolean
End Function
Private Sub ConvertExtSystemDictTypeRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystemDictType)
End Sub
Public Sub CreateExtSystemOrgList()
End Sub
Public Function ExtSystemOrgGetNext(ByRef E As clsExtSystemOrg) As Boolean
End Function
Private Sub ConvertExtSystemOrgRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystemOrg)
End Sub
Public Sub CreateExtSystemPriorityList()
End Sub
Public Function ExtSystemPriorityGetNext(ByRef E As clsExtSystemPriority) As Boolean
End Function
Public Sub DictAuditInsert(Audit As clsDictAudit)
End Sub
Public Sub CreateOrgList()
End Sub
Public Function OrgListGetNext(ByRef Org As clsOrg) As Boolean
End Function
Public Sub CreateDictList(ByVal OrgId As Long, ByVal AuthorId As Long, ByVal TranscriberId As Long, StatusIdStart As Long, StatusIdEnd As Long, TimeStamp As Double, Optional PatId As String = "", Optional PatId2 As String = "")
End Sub
Public Sub CreateDictListFromSQL(SQL As String)
End Sub
Public Function DictListGetNext(ByRef Dict As clsDict) As Boolean
End Function
Public Sub CreateDictTypeList()
End Sub
Public Function DictTypeListGetNext(DictType As clsDictType) As Boolean
End Function
Public Sub CreateOrgDictTypeList()
End Sub
Public Function OrgDictTypeListGetNext(OrgDictType As clsOrgDictType) As Boolean
End Function
Public Sub DeleteOrgDictTypeByOrgId(OrgId As Long)
End Sub
Public Sub SaveOrgDictType(OrgId As Long, DictTypeId As Integer, DictTypeDefault As Boolean)
End Sub
Public Sub CreateOrgPriorityList()
End Sub
Public Function OrgPriorityListGetNext(OrgPriority As clsOrgPriority) As Boolean
End Function
Public Sub DeleteOrgPriorityByOrgId(OrgId As Long)
End Sub
Public Sub SaveOrgPriority(OrgId As Long, PriorityId As Integer, PriorityDefault As Boolean)
End Sub
Public Sub CreateGroupList()
End Sub
Public Function GroupListGetNext(Grp As clsGroup) As Boolean
End Function
Public Sub CreateGroupListForUser(UserId As Long)
End Sub
Public Function GroupListForUserGetNext(ByRef GroupId As Long) As Boolean
End Function
Public Sub CreateUserList()
End Sub
Public Function UserListGetNext(Usr As clsUser) As Boolean
End Function
Public Sub CreatePriorityList()
End Sub
Public Function PriorityListGetNext(Priority As clsPriority) As Boolean
End Function
Public Function DictationStorageOpen(Server As String, _
                                     Database As String, _
                                     UserId As String, _
                                     Password As String) As Boolean
End Function
Public Sub DictationStorageClose()
End Sub
Public Function CheckOutDict(ByRef Dict As clsDict, ByVal DictId As Long, ByVal WithSound As Boolean) As Integer
End Function
Public Function GetDictIdFromExtDictId(ExtSystem As String, ExtDictId As String) As Long
End Function
Public Function CheckInDict(ByRef Dict As clsDict, DiscardChanges As Boolean)
End Function
Public Function CheckInNewDict(ByRef Dict As clsDict) As Long
End Function
Public Sub UnlockDict(DictId As Long)
End Sub
Public Sub CloseRecordset(Rs As ADODB.Recordset)
End Sub
Public Sub AddString(Dest As String, AddPhrase As String, Addendum As String)
End Sub
Public Sub WriteUserData(Ty As String, Ke As String, Va As String)
End Sub
Public Function ReadUserData(Ty As String, Ke As String, Def As String, ByRef version) As String
End Function
Public Sub CreateSysSettingsList(ByVal Ty As String)
End Sub
Public Function SysSettingsListGetNext(ByRef Se As String, ByRef Ke As String, ByRef Va As String) As Boolean
End Function
Public Sub WriteSysSettings(Ty As String, Se As String, Ke As String, s As String)
End Sub
Public Sub WriteStationData(Ty As String, Ke As String, Va As String)
End Sub
Public Function ReadStationData(Ty As String, Ke As String, Def As String, ByRef version) As String
End Function
Public Function SQLDateTimeString(D As Date) As String
End Function
Public Function BuildSQL(SelectFrom As String, WhereClause As String, GroupByClause As String, OrderClause As String)
End Function
