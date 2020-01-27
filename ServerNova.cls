VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServerNova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsServer

'Public Event UIStatusSet(StatusText As String, Busy As Boolean)
'Public Event UIStatusSetSub(SubText As String)
'Public Event UIStatusProgress(Total As Long, Left As Long)
'Public Event UIStatusClear()

Private RsDictList As ADODB.Recordset
Private RsDictTypeList As ADODB.Recordset
Private RsOrgDictTypeList As ADODB.Recordset
Private RsOrgPriorityList As ADODB.Recordset
Private RsGroupList As ADODB.Recordset
Private RsGroupListForUser As ADODB.Recordset
Private RsPriorityList As ADODB.Recordset
Private RsOrgList As ADODB.Recordset
Private RsOrgListClone As ADODB.Recordset
Private RsUserRole As ADODB.Recordset
Private RsUsers As ADODB.Recordset
Private RsStat As ADODB.Recordset
Private RsHist As ADODB.Recordset
Private RsLogg As ADODB.Recordset
Private RsDictAudit As ADODB.Recordset
Private RsDictIdList As ADODB.Recordset
Private RsSysSettingsList As ADODB.Recordset
Private RsExtSystem As ADODB.Recordset
Private RsExtSystemDictType As ADODB.Recordset
Private RsExtSystemOrg As ADODB.Recordset
Private RsExtSystemPriority As ADODB.Recordset

Private DictListCount As Long

Private Const ct_ext_crypto_key = "dfgslkjdfsjdjh843y4jas%&ukj&%hkjdjfg¤&9s8fd"

Private WithEvents Db As clsDbWS
Attribute Db.VB_VarHelpID = -1

Private Sub Class_Initialize()

   Debug.Print "A"
End Sub

Public Property Get clsServer_SQLDateStringFormat() As String

   clsServer_SQLDateStringFormat = Db.SQLDateStringFormat
End Property
Public Property Get clsServer_SQLDateDelimiter() As String

   clsServer_SQLDateDelimiter = Db.SQLDateDelimiter
End Property
Public Property Get clsServer_SQLDateTimeStringFormat() As String

   clsServer_SQLDateTimeStringFormat = Db.SQLDateTimeStringFormat
End Property
Public Property Get clsServer_Server() As String

   clsServer_Server = Db.SQLServer
End Property
Public Property Get clsServer_Database() As String

   clsServer_Database = Db.SQLDatabase
End Property
Public Property Set clsServer_Settings(s As clsSettings)

End Property

Public Property Get clsServer_StorageOpened() As Boolean

   clsServer_StorageOpened = Db.StorageOpened
End Property
Public Function clsServer_IsBatchDoneByClient() As Boolean

   clsServer_IsBatchDoneByClient = False
End Function

Public Sub clsServer_DeleteAllDictations()

   Dim Cmd As New ADODB.Command

   Cmd.CommandTimeout = 500
   OpenStoredProcWithCommand "s_Dictation_DeleteAll", Cmd
End Sub
Public Sub clsServer_DeleteHistory(DeleteOlderThanDays As Integer)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DeleteOlderThanDays)
   
   Cmd.CommandTimeout = 500
   OpenStoredProcWithCommand "s_History_DeleteByCreatedDate", Cmd
End Sub
Public Function clsServer_GetCurrentTimestamp() As Date

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamOutput)
   
   OpenStoredProcWithCommand "s_Dictation_GetCurrentTimestamp", Cmd
   
   clsServer_GetCurrentTimestamp = Cmd(0)
End Function

Private Sub AddHistory(ByRef Dict As clsDict)

   Dim Cmd As New ADODB.Command
   
   Client.Trace.AddRow Trace_Level_Full, "AH", "AHA", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.DictId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.Created))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.OrgId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.DictTypeId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.PriorityId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.ExpiryDate))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.AuthorId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.TranscriberId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.TranscriberOrgId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.TranscribedDate))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.SoundLength)
   
   Client.Trace.AddRow Trace_Level_Full, "AH", "AHB", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   On Error Resume Next
   OpenStoredProcWithCommand "s_History_Insert", Cmd
   Client.Trace.AddRow Trace_Level_Full, "AH", "AHC", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   
End Sub
Private Sub DeleteHistoryByDictId(DictId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
   
   On Error Resume Next
   OpenStoredProcWithCommand "s_History_DeleteByDictId", Cmd
End Sub
Public Function clsServer_GetRoles(Roles As clsRoles, ByVal GroupId As Long, ByVal OrgId As Long) As Boolean

   Dim Rs As ADODB.Recordset
   
   Set Rs = New ADODB.Recordset
   Db.OpenDisconnectedRecordset Rs, "Select * from GroupRoles", "GroupId=" & CStr(GroupId) & " and OrgId=" & CStr(OrgId), "", ""
   
   If Rs.EOF Then
      clsServer_GetRoles = False
   Else
      Set Roles = New clsRoles
      Roles.GroupId = Rs("GroupId")
      Roles.OrgId = Rs("OrgId")
      Roles.Roles = Rs("Roles")
      Roles.Delayed = Rs("Delayed")
      Roles.DelayedHours = Rs("DelayedHours")
      clsServer_GetRoles = True
   End If
   CloseRecordset Rs
End Function
Public Sub clsServer_SaveRoles(Roles As clsRoles)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "GroupRoles"
   Sh.AddFieldMapping "GroupId", Roles.GroupId, fldLong, True
   Sh.AddFieldMapping "OrgId", Roles.OrgId, fldLong, True
   Sh.AddFieldMapping "Roles", Roles.Roles, fldString
   Sh.AddFieldMapping "Delayed", Roles.Delayed, fldBoolean
   Sh.AddFieldMapping "DelayedHours", Roles.DelayedHours, fldLong

   If ExecSQL(Sh.CreateUpdateStmt) <= 0 Then
      ExecSQL Sh.CreateInsertStmt
   End If
   Set Sh = Nothing
End Sub

Public Sub clsServer_SaveOrg(Org As clsOrg)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamReturnValue)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Org.OrgId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Org.OrgParent)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Org.OrgText, 50))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveBool(Org.DictContainer))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveBool(Org.ShowInTree))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveBool(Org.ShowBelow))

   OpenStoredProcWithCommand "s_Org_Set", Cmd
   Org.OrgId = Cmd(0)
End Sub
Public Sub clsServer_SaveDictType(DictType As clsDictType)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "DictType"
   Sh.AddFieldMapping "DictTypeId", DictType.DictTypeId, fldLong, True
   Sh.AddFieldMapping "DictTypeText", DictType.DictTypeText, fldStringNullIfBlank

   If ExecSQL(Sh.CreateUpdateStmt) <= 0 Then
      ExecSQL Sh.CreateInsertStmt
   End If
   Set Sh = Nothing
End Sub
Public Sub clsServer_SaveGroup(Grp As clsGroup)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamReturnValue)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Grp.GroupId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Grp.GroupText, 50))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 255, SaveNullZeroLength(Grp.GroupDesc, 255))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Grp.AdmOrgId))

   OpenStoredProcWithCommand "s_Group_Set", Cmd
   Grp.GroupId = Cmd(0)
End Sub
Public Sub clsServer_SaveUser(Usr As clsUser)

   Dim Sh As New clsSQLHelper
   
   If LCase$(Usr.LoginName) = "sa" Then Exit Sub   'Can't change sa
   
   Sh.Clear
   Sh.Table = "Users"
   Sh.AddFieldMapping "UserId", Usr.UserId, fldLong, True, True
   Sh.AddFieldMapping "LoginName", UCase$(Usr.LoginName), fldString
   Sh.AddFieldMapping "Password", Usr.Password, fldString
   Sh.AddFieldMapping "ShortName", Usr.ShortName, fldString
   Sh.AddFieldMapping "LongName", Usr.LongName, fldString
   Sh.AddFieldMapping "HomeOrgId", Usr.HomeOrgId, fldLongNullIfZero
   Sh.AddFieldMapping "InactivatedByUserId", Usr.InactivatedByUserId, fldDateTimeNullIfZero
   
   If ExecSQL(Sh.CreateUpdateStmt) <= 0 Then
      ExecSQL Sh.CreateInsertStmt
   End If
End Sub
Public Sub clsServer_SaveUserHomeOrg(Usr As clsUser)

   Dim Sh As New clsSQLHelper
   
   If LCase$(Usr.LoginName) = "sa" Then Exit Sub   'Can't change sa
   
   Sh.Clear
   Sh.Table = "Users"
   Sh.AddFieldMapping "UserId", Usr.UserId, fldLong, True
   Sh.AddFieldMapping "HomeOrgId", Usr.HomeOrgId, fldLongNullIfZero
   
   ExecSQL (Sh.CreateUpdateStmt)
End Sub
Public Sub clsServer_ActivateUser(Usr As clsUser, Deactivate As Boolean)

   Dim Sh As New clsSQLHelper
   
   If LCase$(Usr.LoginName) = "sa" Then Exit Sub   'Can't change sa
   
   Sh.Clear
   Sh.Table = "Users"
   Sh.AddFieldMapping "UserId", Usr.UserId, fldLong, True
   If Deactivate Then
      Sh.AddFieldMapping "InactivatedByUserId", Client.User.UserId, fldLong
      Sh.AddFieldMapping "InactivatedTime", Now, fldDateTimeNullIfZero
   Else
      Sh.AddFieldMapping "InactivatedByUserId", 0, fldLongNullIfZero
      Sh.AddFieldMapping "InactivatedTime", 0, fldDateTimeNullIfZero
   End If

   ExecSQL (Sh.CreateUpdateStmt)
End Sub
Public Sub clsServer_DeleteOneUserGroup(UserId As Long, GroupId As Long)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "UserGroup"
   Sh.AddFieldMapping "UserId", UserId, fldLong, True
   Sh.AddFieldMapping "GroupId", GroupId, fldLong, True

   ExecSQL Sh.CreateDeleteStmt
   Set Sh = Nothing

End Sub
Public Sub clsServer_DeleteAllUserGroup(UserId As Long)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "UserGroup"
   Sh.AddFieldMapping "UserId", UserId, fldLong, True

   ExecSQL Sh.CreateDeleteStmt
   Set Sh = Nothing

End Sub
Public Sub clsServer_SaveUserGroup(UserId As Long, GroupId As Long)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "UserGroup"
   Sh.AddFieldMapping "UserId", UserId, fldLong, True
   Sh.AddFieldMapping "GroupId", GroupId, fldLong, True

   ExecSQL Sh.CreateInsertStmt
   Set Sh = Nothing
End Sub
Public Sub clsServer_CreateDictIdList(OrgId As Long, AuthorId As Long, TranscriberId As Long, Optional PatId As String = "", Optional PatId2 As String = "")

   Dim Cmd As New ADODB.Command
   
   If OrgId <> 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , OrgId)
      OpenStoredProcWithRsResultAndCommand RsDictIdList, "s_Dictation_GetIdByOrgId", Cmd
   ElseIf AuthorId <> 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , AuthorId)
      OpenStoredProcWithRsResultAndCommand RsDictIdList, "s_Dictation_GetIdByAuthorId", Cmd
   ElseIf TranscriberId <> 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , TranscriberId)
      OpenStoredProcWithRsResultAndCommand RsDictIdList, "s_Dictation_GetIdByTranscriberId", Cmd
   ElseIf Len(PatId) > 0 Or Len(PatId2) > 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 12, Left$(PatId, 12))
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 12, Left$(PatId2, 12))
      OpenStoredProcWithRsResultAndCommand RsDictIdList, "s_Dictation_GetIdByPatId", Cmd
   Else
      OpenStoredProcWithRsResult RsDictIdList, "s_Dictation_GetIdList"
   End If
End Sub
Public Function clsServer_DictIdGetNext(ByRef DictId As Long) As Boolean

   If RsDictIdList.EOF Then
      clsServer_DictIdGetNext = False
      RsDictIdList.Close
      Set RsDictIdList = Nothing
   Else
      DictId = RsDictIdList("DictId")
      RsDictIdList.MoveNext
      clsServer_DictIdGetNext = True
   End If
End Function
Public Sub clsServer_CreateHist(OrgId As Long, HistYear As Integer, HistType As HistTypeEnum)

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , OrgId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , HistYear)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , HistType)
   
   OpenStoredProcWithRsResultAndCommand RsHist, "s_History_CreateList", Cmd
End Sub
Public Function clsServer_HistGetNext(ByRef Hist As clsHistory) As Boolean

   If RsHist.EOF Then
      clsServer_HistGetNext = False
      RsHist.Close
      Set RsHist = Nothing
   Else
      ConvertHistRsToClass RsHist, Hist
      RsHist.MoveNext
      clsServer_HistGetNext = True
   End If
End Function
Private Sub ConvertHistRsToClass(ByRef Rs As ADODB.Recordset, ByRef Hist As clsHistory)

   Set Hist = New clsHistory
   With Hist
      .Rowid = Rs("RowId")
      .Number(Rs("HistMonth")) = Rs("NumDict")
      .SoundLenSec(Rs("HistMonth")) = Rs("LenSec")
   End With
End Sub
Public Sub clsServer_CreateStat(OrgId As Long, ByRef Stat As clsStat, StatusIdStart As Long, StatusIdEnd As Long, _
                      DaysLimit1 As Integer, _
                      DaysLimit2 As Integer, _
                      DaysLimit3 As Integer, _
                      DaysLimit4 As Integer, _
                      DaysLimit5 As Integer, _
                      DaysLimit6 As Integer, _
                      DaysLimit7 As Integer)
                      
   Dim Cmd As New ADODB.Command
                      
   Set Stat = New clsStat

   Stat.OrgId = OrgId
   Stat.Num1 = 0
   Stat.Num2 = 0
   Stat.Num3 = 0
   Stat.Num4 = 0
   Stat.Num5 = 0
   Stat.Num6 = 0
   Stat.Num7 = 0
   Stat.Num8 = 0
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , OrgId)
   OpenStoredProcWithRsResultAndCommand RsStat, "s_StatNumber", Cmd
   
   Do While Not RsStat.EOF
      If RsStat("StatusId") >= StatusIdStart And RsStat("StatusId") <= StatusIdEnd Then
         If RsStat("DaysLeft") < DaysLimit1 Then
            Stat.Num1 = Stat.Num1 + RsStat("NumberOfDict")
            Stat.LenSec1 = Stat.LenSec1 + RsStat("TotalLength")
         ElseIf RsStat("DaysLeft") < DaysLimit2 Then
            Stat.Num2 = Stat.Num2 + RsStat("NumberOfDict")
            Stat.LenSec2 = Stat.LenSec2 + RsStat("TotalLength")
         ElseIf RsStat("DaysLeft") < DaysLimit3 Then
            Stat.Num3 = Stat.Num3 + RsStat("NumberOfDict")
            Stat.LenSec3 = Stat.LenSec3 + RsStat("TotalLength")
         ElseIf RsStat("DaysLeft") < DaysLimit4 Then
            Stat.Num4 = Stat.Num4 + RsStat("NumberOfDict")
            Stat.LenSec4 = Stat.LenSec4 + RsStat("TotalLength")
         ElseIf RsStat("DaysLeft") < DaysLimit5 Then
            Stat.Num5 = Stat.Num5 + RsStat("NumberOfDict")
            Stat.LenSec5 = Stat.LenSec5 + RsStat("TotalLength")
         ElseIf RsStat("DaysLeft") < DaysLimit6 Then
            Stat.Num6 = Stat.Num6 + RsStat("NumberOfDict")
            Stat.LenSec6 = Stat.LenSec6 + RsStat("TotalLength")
         ElseIf RsStat("DaysLeft") < DaysLimit7 Then
            Stat.Num7 = Stat.Num7 + RsStat("NumberOfDict")
            Stat.LenSec7 = Stat.LenSec7 + RsStat("TotalLength")
         Else
            Stat.Num8 = Stat.Num8 + RsStat("NumberOfDict")
            Stat.LenSec8 = Stat.LenSec8 + RsStat("TotalLength")
         End If
      End If
      RsStat.MoveNext
   Loop
   CloseRecordset RsStat
End Sub
Public Function clsServer_GetSoundfilePath(DictId As Long) As String

   clsServer_GetSoundfilePath = Db.DictationStorageSoundPath & CStr(DictId Mod 10) & "\" & CStr(DictId)
End Function
Public Function clsServer_NeedLoginInfo() As Boolean

   clsServer_NeedLoginInfo = Db.AuthMethod < 2
End Function

Public Function clsServer_UserLogin(User As clsUser, _
                          LoginName As String, _
                          Password As String, _
                          NewPassword As String, _
                          ExtSystem As String, _
                          ExtPassword As String) As Integer

   Dim Rs As New ADODB.Recordset
   Dim Cmd As New ADODB.Command
    
   If Not Db.User_Validate(LoginName, Password, ExtSystem, ExtPassword) Then
      clsServer_UserLogin = 3
      Exit Function
   End If
      
   Client.ExtSystemMgr.Init
      
   If Db.AuthMethod = 2 Then
      If Not SetUserFromSQLLoginName(User) Then
         clsServer_UserLogin = 100
         Exit Function
      End If
      clsServer_UserLogin = 0
      Exit Function
   End If
   
   If Db.AuthMethod < 2 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, UCase(LoginName))
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Password)
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, ExtSystem)
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, RC4(ExtPassword, ct_ext_crypto_key))
   
      OpenStoredProcWithRsResultAndCommand Rs, "s_User_Validate", Cmd
   
      If Not Rs.EOF Then
         If nvl(Rs("UserLockedUntil"), 0) > Now Then
            clsServer_UserLogin = 1
         Else
            If (Len(ExtPassword) > 0 And ExtPassword = Client.ExtSystemMgr.GetExtSystemPassword(ExtSystem)) Or Rs("Password") = Password Then
               User.UserId = Rs("UserId")
               User.LoginName = Rs("LoginName")
               User.ShortName = Rs("ShortName")
               User.LongName = Rs("LongName")
               User.HomeOrgId = nvl(Rs("HomeOrgid"), 0)
               ExecSQL "UPDATE Users SET LastSuccessfulLogin=" & clsServer_SQLDateTimeString(Now) & ",LoginFailureCount=0, UserLockedUntil=NULL WHERE UserId=" & CStr(User.UserId)
               If Len(NewPassword) > 0 Then
                  ExecSQL "UPDATE Users SET Password='" & NewPassword & "' WHERE UserId=" & CStr(User.UserId)
               End If
               clsServer_UserLogin = 0
            Else
               If Rs("LoginFailureCount") >= Client.SysSettings.LoginMaxRetries Then
                  ExecSQL "UPDATE Users SET LoginFailureCount=0,UserLockedUntil=" & clsServer_SQLDateTimeString(DateAdd("n", Client.SysSettings.LoginLockoutMinutes, Now)) & " WHERE LoginName='" & UCase$(LoginName) & "'"
                  clsServer_UserLogin = 1
               Else
                  ExecSQL "UPDATE Users SET LoginFailureCount=LoginFailureCount+1 WHERE LoginName='" & UCase$(LoginName) & "'"
                  clsServer_UserLogin = 2
               End If
            End If
         End If
      Else
         clsServer_UserLogin = 3
      End If
      CloseRecordset Rs
   End If
End Function
Private Function SetUserFromSQLLoginName(User As clsUser) As Boolean

   Dim Rs As ADODB.Recordset

   On Error GoTo SetUserFromSQLLoginName_Err
   
   OpenStoredProcWithRsResult Rs, "s_User_GetCurrentFromSQLLoginName"
   If Rs.EOF Then
      SetUserFromSQLLoginName = False
   Else
      User.UserId = Rs("UserId")
      User.LoginName = Rs("LoginName")
      User.ShortName = Rs("ShortName")
      User.LongName = Rs("LongName")
      User.HomeOrgId = nvl(Rs("HomeOrgid"), 0)
      SetUserFromSQLLoginName = True
   End If
   CloseRecordset Rs
   Exit Function
   
SetUserFromSQLLoginName_Err:
   SetUserFromSQLLoginName = False
   CloseRecordset Rs
   Exit Function
End Function
Private Function ChangeSQLPassword(User As clsUser, OldPassword As String, NewPassword As String)

   ChangeSQLPassword = False
End Function
Public Sub clsServer_CreateLoggList(StartTime As Date, EndTime As Date, StartLevel As Integer, EndLevel As Integer)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(StartTime))
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(EndTime))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StartLevel)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , EndLevel)
   OpenStoredProcWithRsResultAndCommand RsLogg, "s_Logg_GetByDateAndLevel", Cmd
End Sub
Public Function clsServer_LoggListGetNext(ByRef Logg As clsLogg) As Boolean

   If RsLogg.EOF Then
      RsLogg.Close
      Set RsLogg = Nothing
      clsServer_LoggListGetNext = False
      Exit Function
   Else
      ConvertLoggRsToClass RsLogg, Logg
      RsLogg.MoveNext
      clsServer_LoggListGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertLoggRsToClass(ByRef Rs As ADODB.Recordset, ByRef Logg As clsLogg)

   Set Logg = New clsLogg
   With Logg
      .Id = Rs("Id")
      .LoggTime = Rs("LoggTime")
      .LoggId = Rs("LoggId")
      .LoggLevel = Rs("LoggLevel")
      .UserId = nvl(Rs("UserId"), 0)
      .UserShortName = nvl(Rs("UserShortName"), "")
      .StationId = nvl(Rs("StationId"), "")
      .DictId = nvl(Rs("DictId"), 0)
      .LoggData = nvl(Rs("LoggData"), "")
   End With
End Sub
Public Sub clsServer_LoggInsert(Logg As clsLogg)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Logg.LoggId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Logg.LoggLevel)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Logg.UserId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Logg.StationId, 50))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Logg.DictId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Logg.LoggData, 50))
   
   OpenStoredProcWithCommand "s_Logg_Insert", Cmd
End Sub
Public Sub clsServer_CreateDictAuditList(DictId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
   OpenStoredProcWithRsResultAndCommand RsDictAudit, "s_DictAudit_GetByDictId", Cmd
End Sub
Public Function clsServer_DictAuditGetNext(ByRef Audit As clsDictAudit) As Boolean

   If RsDictAudit.EOF Then
      RsDictAudit.Close
      Set RsDictAudit = Nothing
      clsServer_DictAuditGetNext = False
      Exit Function
   Else
      ConvertDictAuditRsToClass RsDictAudit, Audit
      RsDictAudit.MoveNext
      clsServer_DictAuditGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertDictAuditRsToClass(ByRef Rs As ADODB.Recordset, ByRef Audit As clsDictAudit)

   Set Audit = New clsDictAudit
   With Audit
      .Id = Rs("Id")
      .DictId = Rs("DictId")
      .AuditTime = Rs("AuditTime")
      .AuditType = Rs("AuditType")
      .DictStatus = Rs("DictStatus")
      .UserId = nvl(Rs("UserId"), 0)
      .UserShortName = nvl(Rs("UserShortName"), "")
      .StationId = nvl(Rs("StationId"), "")
   End With
End Sub
Public Sub clsServer_CreateDictAuditForUserList(UserId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , UserId)
   OpenStoredProcWithRsResultAndCommand RsDictAudit, "s_DictAudit_GetByUserId", Cmd
End Sub
Public Function clsServer_DictAuditForUserGetNext(ByRef Audit As clsDictAudit) As Boolean

   If RsDictAudit.EOF Then
      RsDictAudit.Close
      Set RsDictAudit = Nothing
      clsServer_DictAuditForUserGetNext = False
      Exit Function
   Else
      ConvertDictAuditForUserRsToClass RsDictAudit, Audit
      RsDictAudit.MoveNext
      clsServer_DictAuditForUserGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertDictAuditForUserRsToClass(ByRef Rs As ADODB.Recordset, ByRef Audit As clsDictAudit)

   Set Audit = New clsDictAudit
   With Audit
      .Id = Rs("Id")
      .DictId = Rs("DictId")
      .AuditTime = Rs("AuditTime")
      .AuditType = Rs("AuditType")
      .DictStatus = Rs("DictStatus")
      .Patient = nvl(Rs("PatId"), "") & "; " & nvl(Rs("PatName"), "")
      .StationId = nvl(Rs("StationId"), "")
   End With
End Sub
Public Sub clsServer_CreateExtSystemList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystem, "s_ExtSystem_Get"
End Sub
Public Function clsServer_ExtSystemGetNext(ByRef E As clsExtSystem) As Boolean

   If RsExtSystem.EOF Then
      RsExtSystem.Close
      Set RsExtSystem = Nothing
      clsServer_ExtSystemGetNext = False
      Exit Function
   Else
      ConvertExtSystemRsToClass RsExtSystem, E
      RsExtSystem.MoveNext
      clsServer_ExtSystemGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertExtSystemRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystem)

   Dim s As String
   
   Set E = New clsExtSystem
   With E
      .ExtSystem = Rs("ExtSystem")
      On Error Resume Next    'Ok if field is not in table
      s = nvl(Rs("ExtSystemPw"), "")
      If Len(s) > 0 Then
         .ExtSystemPw = RC4(s, ct_ext_crypto_key)
      Else
         .ExtSystemPw = ""
      End If
      On Error GoTo 0
   End With
End Sub
Public Sub clsServer_CreateExtSystemDictTypeList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystemDictType, "s_ExtSystemDictType_Get"
End Sub
Public Function clsServer_ExtSystemDictTypeGetNext(ByRef E As clsExtSystemDictType) As Boolean

   If RsExtSystemDictType.EOF Then
      RsExtSystemDictType.Close
      Set RsExtSystemDictType = Nothing
      clsServer_ExtSystemDictTypeGetNext = False
      Exit Function
   Else
      ConvertExtSystemDictTypeRsToClass RsExtSystemDictType, E
      RsExtSystemDictType.MoveNext
      clsServer_ExtSystemDictTypeGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertExtSystemDictTypeRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystemDictType)

   Set E = New clsExtSystemDictType
   With E
      .ExtSystem = Rs("ExtSystem")
      .ExtSystemDictType = Rs("ExtSystemDictType")
      .IntDictTypeId = Rs("IntDictTypeId")
   End With
End Sub
Public Sub clsServer_CreateExtSystemOrgList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystemOrg, "s_ExtSystemOrg_Get"
End Sub
Public Function clsServer_ExtSystemOrgGetNext(ByRef E As clsExtSystemOrg) As Boolean

   If RsExtSystemOrg.EOF Then
      RsExtSystemOrg.Close
      Set RsExtSystemOrg = Nothing
      clsServer_ExtSystemOrgGetNext = False
      Exit Function
   Else
      ConvertExtSystemOrgRsToClass RsExtSystemOrg, E
      RsExtSystemOrg.MoveNext
      clsServer_ExtSystemOrgGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertExtSystemOrgRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystemOrg)

   Set E = New clsExtSystemOrg
   With E
      .ExtSystem = Rs("ExtSystem")
      .ExtSystemOrg = Rs("ExtSystemOrg")
      .IntOrgId = Rs("IntOrgId")
   End With
End Sub
Public Sub clsServer_CreateExtSystemPriorityList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystemPriority, "s_ExtSystemPriority_Get"
End Sub
Public Function clsServer_ExtSystemPriorityGetNext(ByRef E As clsExtSystemPriority) As Boolean

   If RsExtSystemPriority.EOF Then
      RsExtSystemPriority.Close
      Set RsExtSystemPriority = Nothing
      clsServer_ExtSystemPriorityGetNext = False
      Exit Function
   Else
      ConvertExtSystemPriorityRsToClass RsExtSystemPriority, E
      RsExtSystemPriority.MoveNext
      clsServer_ExtSystemPriorityGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertExtSystemPriorityRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystemPriority)

   Set E = New clsExtSystemPriority
   With E
      .ExtSystem = Rs("ExtSystem")
      .ExtSystemPriority = Rs("ExtSystemPriority")
      .IntPriorityId = Rs("IntPriorityId")
   End With
End Sub
Public Sub clsServer_DictAuditInsert(Audit As clsDictAudit)

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Audit.DictId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Audit.AuditType)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Audit.DictStatus)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Audit.UserId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Audit.StationId, 50))
   
   OpenStoredProcWithCommand "s_DictAudit_Insert", Cmd
End Sub
Public Sub clsServer_CreateOrgList()

   Db.OpenDisconnectedRecordset RsOrgList, "SELECT * FROM Org", "", "", ""
   Set RsOrgListClone = RsOrgList.Clone
   Db.OpenDisconnectedRecordset RsUserRole, "SELECT * FROM v_UserRole", "Userid=" & Client.User.UserId, "", ""
   
End Sub
Public Function clsServer_OrgListGetNext(ByRef Org As clsOrg) As Boolean

   Do While True
      If RsOrgList.EOF Then
         RsOrgList.Close
         Set RsOrgList = Nothing
         RsOrgListClone.Close
         Set RsOrgListClone = Nothing
         RsUserRole.Close
         Set RsUserRole = Nothing
         clsServer_OrgListGetNext = False
         Exit Function
      Else
         ConvertOrgRsToClass RsOrgList, Org
         RsOrgList.MoveNext
         If Org.Roles.AnyAllowRoles() Then
            clsServer_OrgListGetNext = True
            Exit Function
         End If
      End If
   Loop
End Function
Private Sub ConvertOrgRsToClass(ByRef Rs As ADODB.Recordset, ByRef Org As clsOrg)

   Dim RolesFromParent As clsRoles
   Dim r As clsRoles
   
   Set Org = New clsOrg
   With Org
      .OrgId = Rs("OrgId")
      .OrgParent = nvl(Rs("OrgParent"), 0)
      .OrgText = Rs("OrgText")
      .ShowInTree = Rs("ShowInTree")
      .DictContainer = Rs("DictContainer")
      .ShowBelow = Rs("ShowBelow")
      Set .Roles = New clsRoles
      GetRolesFromOrgId .Roles, .OrgId
      Set RolesFromParent = New clsRoles
      GetRolesFromOrgParent RolesFromParent, .OrgParent
      AddRole .Roles, RolesFromParent
      If Not RolesFromParent.AnyAllowRoles Then
         .OrgParent = 0
      End If
   End With
End Sub
Private Sub GetRolesFromOrgId(Res As clsRoles, OrgId As Long)

   Dim r As clsRoles
   Dim Sum As clsRoles
   
   If Not RsUserRole.EOF Or Not RsUserRole.BOF Then
      RsUserRole.MoveFirst
      Do While Not RsUserRole.EOF
         If nvl(RsUserRole("Orgid"), 0) = OrgId Then
            Set r = New clsRoles
            
            r.OrgId = RsUserRole("OrgId")
            r.Roles = RsUserRole("Roles")
            r.Delayed = RsUserRole("Delayed")
            r.DelayedHours = RsUserRole("DelayedHours")
            
            AddRole Res, r
         End If
         RsUserRole.MoveNext
      Loop
   End If
End Sub
Private Sub GetRolesFromOrgParent(Res As clsRoles, ByVal Id As Long)

   Dim r As clsRoles
   Dim RParent As clsRoles
   Dim Sum As clsRoles
   
   If Id = 0 Then
   Else
      RsOrgListClone.MoveFirst
      Do While Not RsOrgListClone.EOF
         If Id = RsOrgListClone("OrgId") Then
            GetRolesFromOrgId Res, Id
            GetRolesFromOrgParent Res, nvl(RsOrgListClone("OrgParent"), 0)
            Exit Sub
         End If
         RsOrgListClone.MoveNext
      Loop
   End If
End Sub

Public Sub clsServer_CreateDictList(ByVal OrgId As Long, ByVal AuthorId As Long, ByVal TranscriberId As Long, StatusIdStart As Long, StatusIdEnd As Long, TimeStamp As Double, Optional PatId As String = "", Optional PatId2 As String = "")

   Dim Cmd As New ADODB.Command
   
   DictListCount = 0
   
   If Len(PatId) > 0 Or Len(PatId2) > 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, Left$(PatId, 20))
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, Left$(PatId2, 20))
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdStart)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdEnd)
      
      OpenStoredProcWithRsResultAndCommand RsDictList, "s_Dictation_GetListByPatid2", Cmd
   ElseIf OrgId <> 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , OrgId)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdStart)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdEnd)
      
      OpenStoredProcWithRsResultAndCommand RsDictList, "s_Dictation_GetListByOrgId", Cmd
   ElseIf AuthorId <> 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , AuthorId)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdStart)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdEnd)
      
      OpenStoredProcWithRsResultAndCommand RsDictList, "s_Dictation_GetListByUserId", Cmd
   ElseIf TranscriberId <> 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , TranscriberId)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdStart)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdEnd)
      
      OpenStoredProcWithRsResultAndCommand RsDictList, "s_Dictation_GetListByUserId", Cmd
   Else
      Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , CDate(TimeStamp))
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdStart)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdEnd)
      
      OpenStoredProcWithRsResultAndCommand RsDictList, "s_Dictation_GetListByTimeStamp", Cmd
   End If
End Sub
Public Sub clsServer_CreateDictListFromSQL(SQL As String)

   DictListCount = 0
   Db.OpenRecordset RsDictList, SQL
End Sub

Public Function clsServer_DictListGetNext(ByRef Dict As clsDict) As Boolean

   If RsDictList.EOF Then
      Debug.Print "DictListCount: " & CStr(DictListCount)
      clsServer_DictListGetNext = False
      RsDictList.Close
      Set RsDictList = Nothing
   Else
      DictListCount = DictListCount + 1
      ConvertDictRsToClass RsDictList, Dict
      RsDictList.MoveNext
      clsServer_DictListGetNext = True
   End If
End Function
Private Sub ConvertDictRsToClass(ByRef Rs As ADODB.Recordset, ByRef Dict As clsDict)

   Set Dict = New clsDict
      
   Dict.DictId = Rs("DictId")
   Dict.ExtDictId = nvl(Rs("ExtDictId"), 0)
   Dict.ExtSystem = nvl(Rs("ExtSystem"), "")
   Dict.NoPatient = nvl(Rs("NoPatient"), False)
   Dict.Pat.PatId = nvl(Rs("PatId"), "")
   Dict.Pat.PatId2 = nvl(Rs("PatId2"), "")
   Dict.Pat.PatName = nvl(Rs("PatName"), "")
   Dict.Created = Rs("Created")
   Dict.Changed = nvl(Rs("Changed"), 0)
   Dict.ChangedByUserId = nvl(Rs("ChangedByUserId"), 0)
   Dict.ChangedByUserShortName = nvl(Rs("ChangedByUserShortName"), 0)
   Dict.ChangedByUserLongName = nvl(Rs("ChangedByUserLongName"), 0)
   Dict.SoundLength = nvl(Rs("SoundLength"), 0)
   Dict.OrgId = Rs("OrgId")
   Dict.OrgText = Rs("OrgText")
   Dict.DictTypeId = Rs("DictTypeId")
   Dict.DictTypeText = Rs("DictTypeText")
   Dict.AuthorId = Rs("AuthorId")
   Dict.AuthorShortName = nvl(Rs("AuthorShortName"), "-")
   Dict.AuthorLongName = nvl(Rs("AuthorLongName"), "-")
   Dict.TranscriberId = nvl(Rs("TranscriberId"), 0)
   Dict.TranscriberOrgId = nvl(Rs("TranscriberOrgId"), 0)
   Dict.TranscribedDate = nvl(Rs("TranscribedDate"), 0)
   Dict.TranscriberShortName = nvl(Rs("TranscriberShortName"), "")
   Dict.TranscriberLongName = nvl(Rs("TranscriberLongName"), "")
   Dict.StatusId = Rs("StatusId")
   Dict.StatusText = Rs("StatusText")
   Dict.PriorityId = Rs("PriorityId")
   Dict.PriorityText = Rs("PriorityText")
   Dict.ExpiryDate = Rs("ExpiryDate")
   Dict.Txt = nvl(Rs("Txt"), "")
   Dict.LockedByStation = nvl(Rs("LockedByStation"), "")
   Dict.LockedByUserShortName = nvl(Rs("LockedByUserShortName"), "")
   Dict.LockedByUserLongName = nvl(Rs("LockedByUserLongName"), "")
   Dict.LockedTime = nvl(Rs("LockedTime"), 0)
   Dict.TimeStamp = Rs("Timestamp")
   Dict.SoundDeleted = nvl(Rs("SoundDeleted"), 0)
   Dict.Note = nvl(Rs("Note"), "")
   Dict.AuthorPos = nvl(Rs("AuthorPos"), 0)
   Dict.TranscriberPos = nvl(Rs("TranscriberPos"), 0)
   
   Dict.SoundDirty = False
   Dict.InfoDirty = False
End Sub
Public Sub clsServer_CreateDictTypeList()

   OpenStoredProcWithRsResult RsDictTypeList, "s_DictType_Get"
End Sub
Public Function clsServer_DictTypeListGetNext(DictType As clsDictType) As Boolean

   If RsDictTypeList.EOF Then
      clsServer_DictTypeListGetNext = False
      RsDictTypeList.Close
      Set RsDictTypeList = Nothing
   Else
      ConvertDictTypeRsToClass RsDictTypeList, DictType
      RsDictTypeList.MoveNext
      clsServer_DictTypeListGetNext = True
   End If
End Function
Private Sub ConvertDictTypeRsToClass(ByRef Rs As ADODB.Recordset, ByRef DictType As clsDictType)

   Set DictType = New clsDictType

   DictType.DictTypeId = Rs("DictTypeId")
   DictType.DictTypeText = Rs("DictTypeText")
End Sub
Public Sub clsServer_CreateOrgDictTypeList()

   OpenStoredProcWithRsResult RsOrgDictTypeList, "s_OrgDictType_Get"
End Sub
Public Function clsServer_OrgDictTypeListGetNext(OrgDictType As clsOrgDictType) As Boolean

   If RsOrgDictTypeList.EOF Then
      clsServer_OrgDictTypeListGetNext = False
      RsOrgDictTypeList.Close
      Set RsOrgDictTypeList = Nothing
   Else
      ConvertOrgDictTypeRsToClass RsOrgDictTypeList, OrgDictType
      RsOrgDictTypeList.MoveNext
      clsServer_OrgDictTypeListGetNext = True
   End If
End Function
Private Sub ConvertOrgDictTypeRsToClass(ByRef Rs As ADODB.Recordset, ByRef OrgDictType As clsOrgDictType)

   Set OrgDictType = New clsOrgDictType

   OrgDictType.OrgId = Rs("OrgId")
   OrgDictType.DictTypeId = Rs("DictTypeId")
   OrgDictType.Def = Rs("Def") <> 0
End Sub
Public Sub clsServer_DeleteOrgDictTypeByOrgId(OrgId As Long)

   Dim Sh As New clsSQLHelper

   Sh.Clear
   Sh.Table = "OrgDictType"
   Sh.AddFieldMapping "OrgId", OrgId, fldLong, True
   
   ExecSQL Sh.CreateDeleteStmt
   Set Sh = Nothing
End Sub
Public Sub clsServer_SaveOrgDictType(OrgId As Long, DictTypeId As Integer, DictTypeDefault As Boolean)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "OrgDictType"
   Sh.AddFieldMapping "OrgId", OrgId, fldLong, True
   Sh.AddFieldMapping "DictTypeId", DictTypeId, fldLong
   Sh.AddFieldMapping "Def", DictTypeDefault, fldBoolean

   ExecSQL Sh.CreateInsertStmt
   Set Sh = Nothing
End Sub
Public Sub clsServer_CreateOrgPriorityList()

   OpenStoredProcWithRsResult RsOrgPriorityList, "s_OrgPriority_Get"
End Sub
Public Function clsServer_OrgPriorityListGetNext(OrgPriority As clsOrgPriority) As Boolean

   If RsOrgPriorityList.EOF Then
      clsServer_OrgPriorityListGetNext = False
      RsOrgPriorityList.Close
      Set RsOrgPriorityList = Nothing
   Else
      ConvertOrgPriorityRsToClass RsOrgPriorityList, OrgPriority
      RsOrgPriorityList.MoveNext
      clsServer_OrgPriorityListGetNext = True
   End If
End Function
Private Sub ConvertOrgPriorityRsToClass(ByRef Rs As ADODB.Recordset, ByRef OrgPriority As clsOrgPriority)

   Set OrgPriority = New clsOrgPriority

   OrgPriority.OrgId = Rs("OrgId")
   OrgPriority.PriorityId = Rs("PriorityId")
   OrgPriority.Def = Rs("Def") <> 0
End Sub
Public Sub clsServer_DeleteOrgPriorityByOrgId(OrgId As Long)

   Dim Sh As New clsSQLHelper

   Sh.Clear
   Sh.Table = "OrgPriority"
   Sh.AddFieldMapping "OrgId", OrgId, fldLong, True
   
   ExecSQL Sh.CreateDeleteStmt
   Set Sh = Nothing
End Sub
Public Sub clsServer_SaveOrgPriority(OrgId As Long, PriorityId As Integer, PriorityDefault As Boolean)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "OrgPriority"
   Sh.AddFieldMapping "OrgId", OrgId, fldLong, True
   Sh.AddFieldMapping "PriorityId", PriorityId, fldInteger
   Sh.AddFieldMapping "Def", PriorityDefault, fldBoolean

   ExecSQL Sh.CreateInsertStmt
   Set Sh = Nothing
End Sub
Public Sub clsServer_CreateGroupList()

   OpenStoredProcWithRsResult RsGroupList, "s_Group_Get"
End Sub
Public Function clsServer_GroupListGetNext(Grp As clsGroup) As Boolean

   If RsGroupList.EOF Then
      clsServer_GroupListGetNext = False
      RsGroupList.Close
      Set RsGroupList = Nothing
   Else
      ConvertGroupRsToClass RsGroupList, Grp
      RsGroupList.MoveNext
      clsServer_GroupListGetNext = True
   End If
End Function
Private Sub ConvertGroupRsToClass(ByRef Rs As ADODB.Recordset, ByRef Grp As clsGroup)

   Set Grp = New clsGroup

   Grp.GroupId = Rs("GroupId")
   Grp.GroupText = Rs("GroupText")
   Grp.GroupDesc = nvl(Rs("GroupDesc"), "")
   Grp.AdmOrgId = nvl(Rs("AdmOrgId"), 0)
End Sub
Public Sub clsServer_CreateGroupListForUser(UserId As Long)

   Db.OpenDisconnectedRecordset RsGroupListForUser, "SELECT * FROM UserGroup", "Userid=" & CStr(UserId), "", ""
End Sub
Public Function clsServer_GroupListForUserGetNext(ByRef GroupId As Long) As Boolean

   If RsGroupListForUser.EOF Then
      clsServer_GroupListForUserGetNext = False
      RsGroupListForUser.Close
      Set RsGroupListForUser = Nothing
   Else
      GroupId = RsGroupListForUser("GroupId")
      RsGroupListForUser.MoveNext
      clsServer_GroupListForUserGetNext = True
   End If
End Function
Public Sub clsServer_CreateUserList()

   Db.OpenDisconnectedRecordset RsUsers, "SELECT * FROM Users", "", "", "ShortName ASC"
End Sub
Public Function clsServer_UserListGetNext(Usr As clsUser) As Boolean

   If RsUsers.EOF Then
      clsServer_UserListGetNext = False
      CloseRecordset RsUsers
   Else
      ConvertUsersRsToClass RsUsers, Usr
      RsUsers.MoveNext
      clsServer_UserListGetNext = True
   End If
End Function
Private Sub ConvertUsersRsToClass(ByRef Rs As ADODB.Recordset, ByRef Usr As clsUser)

   Set Usr = New clsUser

   Usr.UserId = Rs("UserId")
   Usr.LoginName = Rs("LoginName")
   Usr.Password = Rs("Password")
   Usr.ShortName = Rs("ShortName")
   Usr.LongName = Rs("LongName")
   Usr.HomeOrgId = nvl(Rs("HomeOrgId"), 0)
   Usr.InactivatedByUserId = nvl(Rs("InactivatedByUserId"), 0)
   Usr.InactivatedTime = nvl(Rs("InactivatedTime"), 0)
End Sub
Public Sub clsServer_CreatePriorityList()

   Db.OpenDisconnectedRecordset RsPriorityList, "SELECT * FROM Priority", "", "", "PriorityText ASC"
End Sub
Public Function clsServer_PriorityListGetNext(Priority As clsPriority) As Boolean

   If RsPriorityList.EOF Then
      clsServer_PriorityListGetNext = False
      RsPriorityList.Close
      Set RsPriorityList = Nothing
   Else
      ConvertPriorityRsToClass RsPriorityList, Priority
      RsPriorityList.MoveNext
      clsServer_PriorityListGetNext = True
   End If
End Function
Private Sub ConvertPriorityRsToClass(ByRef Rs As ADODB.Recordset, ByRef Priority As clsPriority)

   Set Priority = New clsPriority
   
   Priority.PriorityId = Rs("PriorityId")
   Priority.PriortyText = Rs("PriorityText")
   Priority.Days = Rs("Days")
   Priority.Warning = nvl(Rs("Warning"), 0) <> 0
   Priority.Colour = nvl(Rs("Colour"), 0)
End Sub

Private Sub Class_Terminate()

   clsServer_DictationStorageClose
   Set Db = Nothing
End Sub

Public Function clsServer_DictationStorageOpen(Server As String, _
                                     Database As String, _
                                     UserId As String, _
                                     Password As String) As Boolean

   Set Db = New clsDbWS
   clsServer_DictationStorageOpen = Db.DictationStorageOpen(Server, Database, UserId, Password)
End Function

Public Sub clsServer_DictationStorageClose()

   Db.DictationStorageClose
End Sub

Public Function clsServer_CheckOutDict(ByRef Dict As clsDict, ByVal DictId As Long, ByVal WithSound As Boolean) As Integer

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   'RaiseEvent UIStatusSet(Client.Texts.Txt(1260101, "Diktatet hämtas"), True)
   Client.UIStatusSet Client.Texts.Txt(1260101, "Diktatet hämtas"), True
   If LockDictation(DictId) Then
      
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
        
      OpenStoredProcWithRsResultAndCommand Rs, "s_Dictation_GetById", Cmd
   
      If Rs.EOF Then
         CloseRecordset Rs
         clsServer_CheckOutDict = 2
         GoTo CheckOutDict_Exit
      End If
      
      ConvertDictRsToClass Rs, Dict
      
      If Dict.AuthorId = Client.User.UserId Then
         Dict.CurrentPos = Dict.AuthorPos
      Else
         If Dict.TranscriberId = Client.User.UserId Then
            Dict.CurrentPos = Dict.TranscriberPos
         End If
      End If
      
      If WithSound And Dict.StatusId < SoundDeleted Then
         CheckOutSound Dict
      End If
     
      clsServer_CheckOutDict = 0
   Else
      clsServer_CheckOutDict = 1
   End If
CheckOutDict_Exit:
   'RaiseEvent UIStatusClear
   Client.UIStatusClear
End Function
Private Sub CheckOutSound(Dict As clsDict)

   If CheckoutSoundValues(Dict) Then
      Client.DictFileMgr.CreateLocalValues Dict.LocalDictFile
      CopyFileFromSoundStorage Dict.DictId, Dict.LocalDictFile.LocalFilenameFormat
      Client.DictFileMgr.ConvertOrgFormatToPlayFormat Dict.LocalDictFile
   End If
End Sub
Private Function CheckoutSoundValues(Dict As clsDict) As Boolean

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Res As Boolean
   
   SQL = clsServer_BuildSQL("Select FileName, SoundType from Sound", "DictId=" & CStr(Dict.DictId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   Db.OpenRecordset Rs, SQL
      
   If Not Rs.EOF Then
      Dict.LocalDictFile.OrgFileName = nvl(Rs("FileName"), "")
      Dict.LocalDictFile.OrgType = nvl(Rs("SoundType"), "")
      Res = True
   Else
      Res = False
   End If
   
   CloseRecordset Rs

   CheckoutSoundValues = Res
End Function
Public Function clsServer_GetDictIdFromExtDictId(ExtSystem As String, ExtDictId As String) As Long

   If Len(ExtSystem) = 0 Or Len(ExtDictId) = 0 Then Exit Function
   
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamReturnValue)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, SaveNullZeroLength(ExtSystem, 10))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, SaveNullZeroLength(ExtDictId, 20))

   OpenStoredProcWithCommand "s_Dictation_GetDictIdFromExtDictId", Cmd
   
   clsServer_GetDictIdFromExtDictId = Cmd(0)
   
End Function
Private Function LockDictation(DictId As Long) As Boolean

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamReturnValue)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Client.User.UserId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Client.Station.Id, 50))

   OpenStoredProcWithCommand "s_Dictation_LockByUser", Cmd
   
   LockDictation = (Cmd(0) = 0)
End Function
Private Sub UnLockDictationWithNewStatus(DictId As Long, StatusId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusId)

   OpenStoredProcWithCommand "s_Dictation_UnLockWithNewStatus", Cmd
End Sub
Public Function clsServer_CheckInDict(ByRef Dict As clsDict, DiscardChanges As Boolean)

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   Dim Dirty As Boolean
   
   Client.Trace.AddRow Trace_Level_Full, "CID", "CIDA", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   'RaiseEvent UIStatusSet(Client.Texts.Txt(1260102, "Diktatet sparas"), True)
   Client.UIStatusSet Client.Texts.Txt(1260102, "Diktatet sparas"), True
      
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.DictId)
   
   Client.Trace.AddRow Trace_Level_Full, "CID", "CIDB", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   OpenStoredProcWithRsResultAndCommand Rs, "s_Dictation_GetById", Cmd
   Client.Trace.AddRow Trace_Level_Full, "CID", "CIDC", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   
   If Rs.EOF Then
      Client.Trace.AddRow Trace_Level_Full, "CID", "CIDD", "", CStr(Dict.DictId), CStr(Dict.StatusId)
      clsServer_CheckInDict = 1
   Else
      Client.Trace.AddRow Trace_Level_Full, "CID", "CIDE", "", CStr(Dict.DictId), CStr(Dict.StatusId)
      If Rs("LockedByUserId") <> Client.User.UserId Then
         Client.Trace.AddRow Trace_Level_Full, "CID", "CIDF", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         clsServer_CheckInDict = 2
      Else
         Client.Trace.AddRow Trace_Level_Full, "CID", "CIDG", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         If Not DiscardChanges Then
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDH", "", CStr(Dict.DictId), CStr(Dict.StatusId)
            'If Rs("AuthorId") = Client.User.UserId Then
               If Rs("NoPatient") <> Dict.NoPatient Then
                  Dirty = True
               End If
               If Rs("PatId") <> Dict.Pat.PatId Then
                  Dirty = True
               End If
               If Rs("PatId2") <> Dict.Pat.PatId2 Then
                  Dirty = True
               End If
               If Rs("PatName") <> Dict.Pat.PatName Then
                  Dirty = True
               End If
               If Rs("OrgId") <> Dict.OrgId Then
                  Dirty = True
               End If
               If Rs("DictTypeId") <> Dict.DictTypeId Then
                  Dirty = True
               End If
               If Rs("PriorityId") <> Dict.PriorityId Then
                  Dirty = True
               End If
               If Rs("ExpiryDate") <> Dict.ExpiryDate Then
                  Dirty = True
               End If
               If nvl(Rs("Txt"), "") <> Dict.Txt Then
                  Dirty = True
               End If
               If nvl(Rs("Note"), "") <> Dict.Note Then
                  Dirty = True
               End If
               
               If Dict.LocalDictFile.IsSoundData Then
                  If Dict.SoundDirty Then
                     CheckInSound Dict
                  End If
               End If
               
               If Dirty Or Dict.SoundDirty Then
                  Dict.Changed = Now
                  Dict.ChangedByUserId = Client.User.UserId
               End If
            'End If
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDI", "", CStr(Dict.DictId), CStr(Dict.StatusId)
            If Rs("StatusId") > Recorded And Dict.StatusId <= Recorded Then
               DeleteHistoryByDictId Dict.DictId
               Dict.TranscribedDate = 0
               Dict.TranscriberId = 0
               Dict.TranscriberOrgId = 0
            End If
            If Rs("StatusId") < BeingTrancribed And Dict.StatusId >= BeingTrancribed Then
               Dict.TranscriberId = Client.User.UserId
               If Client.User.HomeOrgId > 0 Then
                  Dict.TranscriberOrgId = Client.User.HomeOrgId
               End If
            End If
            If Rs("StatusId") < WaitForSign And (Dict.StatusId = WaitForSign Or Dict.StatusId = Transcribed) Then
               Dict.TranscriberId = Client.User.UserId
               If Client.User.HomeOrgId > 0 Then
                  Dict.TranscriberOrgId = Client.User.HomeOrgId
               End If
               Dict.TranscribedDate = Now
               Client.Trace.AddRow Trace_Level_Full, "CID", "CIDJ", "", CStr(Dict.DictId), CStr(Dict.StatusId)
               AddHistory Dict
               Client.Trace.AddRow Trace_Level_Full, "CID", "CIDK", "", CStr(Dict.DictId), CStr(Dict.StatusId)
            End If
            
            If Dict.AuthorId = Client.User.UserId Then
               Dict.AuthorPos = Dict.CurrentPos
            Else
               If Dict.TranscriberId = Client.User.UserId Then
                  Dict.TranscriberPos = Dict.CurrentPos
               End If
            End If
            
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDL", "", CStr(Dict.DictId), CStr(Dict.StatusId)
            clsServer_CheckInDict = CheckInDictInternal(Dict)
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDM", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         Else
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDN", "", CStr(Dict.DictId), CStr(Dict.StatusId)
            UnLockDictationWithNewStatus Dict.DictId, Dict.StatusId
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDO", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         End If
      End If
   End If
   CloseRecordset Rs
   'RaiseEvent UIStatusClear
   Client.UIStatusClear
End Function
Private Sub CheckInSound(Dict As clsDict)

   CopyFileToSoundStorage Dict.LocalDictFile.LocalFilenameFormat, Dict.DictId, Dict.LocalDictFile.OrgFileName, Dict.LocalDictFile.OrgType
End Sub
Private Function CheckInDictInternal(ByRef Dict As clsDict) As Integer

   Dim Cmd As New ADODB.Command
   Dim Dirty As Boolean
   
   Client.Trace.AddRow Trace_Level_Full, "CIDI", "CIDIA", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamReturnValue)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.DictId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, SaveNullZeroLength(Dict.ExtDictId, 20))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, SaveNullZeroLength(Dict.ExtSystem, 10))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveBool(Dict.NoPatient))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, SaveNullZeroLength(Dict.Pat.PatId, 20))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, SaveNullZeroLength(Dict.Pat.PatId2, 20))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Dict.Pat.PatName, 50))
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.Changed))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Dict.ChangedByUserId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.OrgId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.DictTypeId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.AuthorId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Dict.TranscriberId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Dict.TranscriberOrgId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.TranscribedDate))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.StatusId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.PriorityId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.ExpiryDate))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Dict.Txt, 50))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.SoundLength)
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(Dict.SoundDeleted))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 100, SaveNullZeroLength(Dict.Note, 100))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.AuthorPos)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.TranscriberPos)

   Client.Trace.AddRow Trace_Level_Full, "CIDI", "CIDIB", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   OpenStoredProcWithCommand "s_Dictation_CheckIn3", Cmd
   Client.Trace.AddRow Trace_Level_Full, "CIDI", "CIDIC", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   Dict.DictId = Cmd(0)
   Client.Trace.AddRow Trace_Level_Full, "CIDI", "CIDID", "", CStr(Dict.DictId), CStr(Dict.StatusId)
End Function
Public Function clsServer_CheckInNewDict(ByRef Dict As clsDict) As Long

   'RaiseEvent UIStatusSet(Client.Texts.Txt(1260102, "Diktatet sparas"), True)
   Client.UIStatusSet Client.Texts.Txt(1260102, "Diktatet sparas"), True
      Dict.AuthorId = Client.User.UserId
      Dict.AuthorPos = Dict.CurrentPos
      CheckInDictInternal Dict
      
      If Dict.LocalDictFile.IsSoundData Then
         CheckInSound Dict
      End If
      clsServer_CheckInNewDict = Dict.DictId
   'RaiseEvent UIStatusClear
   Client.UIStatusClear
End Function
Public Sub clsServer_UnlockDict(DictId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)

   OpenStoredProcWithCommand "s_Dictation_Unlock", Cmd
   
End Sub

Private Sub AddRole(SumRoles As clsRoles, ByVal Roles2 As clsRoles)

   Dim I As Integer
   Dim Letter As String
   Dim UcaseLetter As String
   Dim UcaseSumRoles As String
   Dim R2 As String
   Dim DoWeAlreadyHaveARole As Boolean
   
   SumRoles.GroupId = Roles2.GroupId
   SumRoles.OrgId = Roles2.OrgId
   
   DoWeAlreadyHaveARole = Len(SumRoles.Roles) > 0
   
   R2 = Roles2.Roles
   For I = 1 To Len(R2)
      Letter = mId$(R2, I, 1)
      UcaseLetter = UCase(Letter)
      UcaseSumRoles = UCase(SumRoles.Roles)
      If InStr(UcaseSumRoles, UcaseLetter) = 0 Then
         SumRoles.Roles = SumRoles.Roles & Letter
      End If
   Next I
   
   If DoWeAlreadyHaveARole Then
      SumRoles.Delayed = SumRoles.Delayed And Roles2.Delayed
      If SumRoles.Delayed Then
         If Roles2.DelayedHours < SumRoles.DelayedHours Then
            SumRoles.DelayedHours = Roles2.DelayedHours
         End If
      Else
         SumRoles.DelayedHours = 0
      End If
   Else
      SumRoles.Delayed = Roles2.Delayed
      SumRoles.DelayedHours = Roles2.DelayedHours
   End If
End Sub

Private Sub CloseRecordset(Rs As ADODB.Recordset)

   clsServer_CloseRecordset Rs
End Sub

Public Sub clsServer_CloseRecordset(Rs As ADODB.Recordset)

   On Error Resume Next
   Rs.Close
   Set Rs = Nothing
End Sub

Private Sub OpenStoredProcWithRsResult(ByRef Rs As ADODB.Recordset, SpName As String)

   Dim Cmd As ADODB.Command
   
   Set Cmd = New ADODB.Command
   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   Db.ExecuteCommandWithRsResult Rs, Cmd
End Sub
Private Sub OpenStoredProcWithRsResultAndCommand(ByRef Rs As ADODB.Recordset, SpName As String, ByRef Cmd As ADODB.Command)

   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   Db.ExecuteCommandWithRsResult Rs, Cmd
End Sub
Private Sub OpenStoredProcWithCommand(SpName As String, ByRef Cmd As ADODB.Command)

   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   Db.ExecuteCommand Cmd
End Sub


Public Sub clsServer_AddString(Dest As String, AddPhrase As String, Addendum As String)

   If Len(Addendum) = 0 Then Exit Sub
   If Len(Dest) = 0 Then
      Dest = Addendum
   Else
      Dest = Dest & " " & AddPhrase & " " & Addendum
   End If
End Sub
Public Sub clsServer_WriteUserData(Ty As String, Ke As String, Va As String)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Client.User.UserId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, Ty)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, Ke)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 4000, Va)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, ApplicationVersion)
   
   OpenStoredProcWithCommand "s_UserData_Set", Cmd

End Sub
Public Function clsServer_ReadUserData(Ty As String, Ke As String, Def As String, ByRef version) As String

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Client.User.UserId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, Ty)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, Ke)
   
   OpenStoredProcWithRsResultAndCommand Rs, "s_UserData_Get", Cmd
   
   If Not Rs.EOF Then
      clsServer_ReadUserData = Rs("Va")
      version = Rs("Version")
   End If
   CloseRecordset Rs
End Function
Public Sub clsServer_CreateSysSettingsList(ByVal Ty As String)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Ty)
   
   OpenStoredProcWithRsResultAndCommand RsSysSettingsList, "s_SysSettings_GetByTy", Cmd
End Sub
Public Function clsServer_SysSettingsListGetNext(ByRef Se As String, ByRef Ke As String, ByRef Va As String) As Boolean

   If RsSysSettingsList.EOF Then
      clsServer_SysSettingsListGetNext = False
      RsSysSettingsList.Close
      Set RsSysSettingsList = Nothing
   Else
      Se = RsSysSettingsList("Se")
      Ke = RsSysSettingsList("Ke")
      Va = RsSysSettingsList("Va")
      RsSysSettingsList.MoveNext
      clsServer_SysSettingsListGetNext = True
   End If
End Function
Public Sub clsServer_WriteSysSettings(Ty As String, Se As String, Ke As String, s As String)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Ty)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Se)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Ke)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 200, s)
   
   OpenStoredProcWithCommand "s_SysSettings_Set", Cmd
   
End Sub
Public Sub clsServer_WriteStationData(Ty As String, Ke As String, Va As String)

   Dim Sh As New clsSQLHelper
   
   Sh.Clear
   Sh.Table = "StationData"
   Sh.AddFieldMapping "Ty", Ty, fldString, True
   Sh.AddFieldMapping "Ke", Ke, fldString, True
   Sh.AddFieldMapping "StationId", Client.Station.Id, fldString, True
   Sh.AddFieldMapping "Va", Va, fldString
   Sh.AddFieldMapping "Version", ApplicationVersion, fldString

   If ExecSQL(Sh.CreateUpdateStmt) <= 0 Then
      ExecSQL Sh.CreateInsertStmt
   End If
   Set Sh = Nothing
End Sub
Public Function clsServer_ReadStationData(Ty As String, Ke As String, Def As String, ByRef version) As String

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = clsServer_BuildSQL("Select * from StationData", "Ty='" & Ty & "' and Ke='" & Ke & "' and StationId='" & Client.Station.Id & "'", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   Db.OpenRecordset Rs, SQL
   If Rs.EOF Then
      clsServer_ReadStationData = Def
      version = ""
   Else
      clsServer_ReadStationData = Rs("Va")
      version = Rs("Version")
   End If
   CloseRecordset Rs
End Function

Private Sub CopyFileToSoundStorage(SourcePath As String, DictId As Long, FileName As String, SoundType As String)

   Dim SoundPath As String

   'RaiseEvent UIStatusSetSub(Client.Texts.Txt(1260103, "Ljudet sparas..."))
   Client.UIStatusSetSub Client.Texts.Txt(1260103, "Ljudet sparas...")
   Db.CopyFileToSoundTable SourcePath, DictId, FileName, SoundType
   'RaiseEvent UIStatusSetSub("")
   Client.UIStatusSetSub ""
End Sub

Private Sub CopyFileFromSoundStorage(DictId As Long, DestPath As String)

   Dim SoundPath As String

   'RaiseEvent UIStatusSetSub(Client.Texts.Txt(1260104, "Ljudet hämtas..."))
   Client.UIStatusSetSub Client.Texts.Txt(1260104, "Ljudet hämtas...")
   If Client.SysSettings.SoundUseImageField Then
      Db.CopyFileFromSoundTable DictId, DestPath
   Else
      SoundPath = clsServer_GetSoundfilePath(DictId)
      FileCopyWithRetry SoundPath, DestPath
   End If
   'RaiseEvent UIStatusSetSub("")
   Client.UIStatusSetSub ""
End Sub
Private Sub FileCopyWithRetry(Source As String, Destination As String)

   Dim Retries As Integer
   Dim ErrVal As Long
   
   Retries = 3
   On Error GoTo FileCopyWithRetry_Err
   FileCopy Source, Destination
   Exit Sub
   
FileCopyWithRetry_Err:
   If Retries > 0 Then
      Retries = Retries - 1
      Resume
   Else
      Error ErrVal
   End If
End Sub


Private Sub CreateSoundSubFolders()
  
   Dim I As Integer
   
   On Error Resume Next
   For I = 0 To 9
      MkDir Db.DictationStorageSoundPath & CStr(I)
      SetAttr Db.DictationStorageSoundPath & CStr(I), vbReadOnly + vbHidden
   Next I
End Sub
Private Function SQLDateString(D As Date) As String

   SQLDateString = clsServer_SQLDateDelimiter & Format$(D, clsServer_SQLDateStringFormat) & clsServer_SQLDateDelimiter
End Function
Public Function clsServer_SQLDateTimeString(D As Date) As String

   clsServer_SQLDateTimeString = clsServer_SQLDateDelimiter & Format$(D, clsServer_SQLDateStringFormat) & clsServer_SQLDateDelimiter
End Function

Private Function SaveBool(B As Boolean) As Integer

   If B Then
      SaveBool = 1
   Else
      SaveBool = 0
   End If
End Function

Private Function SaveDate(D As Date) As Variant

   If D = 0 Then
      SaveDate = Null
   Else
      SaveDate = D
   End If
End Function
Private Function SaveNullIfZero(Value As Long) As Variant

   If Value = 0 Then
      SaveNullIfZero = Null
   Else
      SaveNullIfZero = Value
   End If
End Function
Private Function SaveNullZeroLength(s As String, MAXLEN As Integer) As Variant

   If Len(s) = 0 Then
      SaveNullZeroLength = Null
   Else
      SaveNullZeroLength = Left$(s, MAXLEN)
   End If
End Function
Public Function clsServer_BuildSQL(SelectFrom As String, WhereClause As String, GroupByClause As String, OrderClause As String)

   clsServer_BuildSQL = Db.BuildSQL(SelectFrom, WhereClause, GroupByClause, OrderClause)
End Function
Private Function ExecSQL(SQL As String) As Long

   Dim Cmd As New ADODB.Command
   
   Cmd.CommandText = SQL
   Cmd.CommandType = adCmdText
   
   ExecSQL = Db.ExecuteCommand(Cmd)
End Function

Private Sub Db_UIStatusClear()

   'RaiseEvent UIStatusClear
   Client.UIStatusClear
End Sub

Private Sub Db_UIStatusProgress(Total As Long, Left As Long)

   'RaiseEvent UIStatusProgress(Total, Left)
   Client.UIStatusProgress Total, Left
End Sub

Private Sub Db_UIStatusSet(StatusText As String, Busy As Boolean)

   'RaiseEvent UIStatusSet(StatusText, Busy)
   Client.UIStatusSet StatusText, Busy
End Sub

Private Sub Db_UIStatusSetSub(SubText As String)

   'RaiseEvent UIStatusSetSub(SubText)
   Client.UIStatusSetSub SubText
End Sub
