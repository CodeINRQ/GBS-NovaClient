VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServerSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event UIStatusSet(StatusText As String, Busy As Boolean)
Public Event UIStatusSetSub(SubText As String)
Public Event UIStatusProgress(Total As Long, Left As Long)
Public Event UIStatusClear()

Private Const DBUserid As String = "ct_public"
Private Const DBPassword As String = "ssdjgfUhssdF&&ghdsj3kjek34094fkkkkkk"
Private Const ct_app_default_password = "kjsd9sd9873hewls983939899%%3_98kdfjjknvy111"
Private Const ct_app_crypto_key = "dkllsahd84h5uig78djdslkf5%%7udfsn#nkgkk"
Private Const ct_ext_crypto_key = "dfgslkjdfsjdjh843y4jas%&ukj&%hkjdjfg¤&9s8fd"
'Private Const DBUseridAdmin As String = "CTAdmin"
'Private Const DBPasswordAdmin As String = "ssdjgfUhssdFkdhh555"

Private Settings As clsSettings

Private DictationStorageSoundPath As String
Private DictationStoragePath As String
Public SQLDateStringFormat As String
Public SQLDateDelimiter As String

Private mConnectionString As String
Private mStorageOpened As Boolean               'True is Storage opened (connections may be temp closed)
Private mAuthMethod As Integer                  '0 ct_user, 1 SQL user, 2 Trusted connection
Private mCryptoAppRolePassword As String        'Encrypted password for app role ct_app

Private Db As ADODB.Connection
Private SQLServer As String
Private SQLDatabase As String

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
Public Property Get Server() As String

   Server = SQLServer
End Property
Public Property Get Database() As String

   Database = SQLDatabase
End Property
Public Property Get StorageOpened() As Boolean

   StorageOpened = mStorageOpened
End Property
Public Function IsBatchDoneByClient() As Boolean

   IsBatchDoneByClient = False
End Function

Public Sub DeleteAllDictations()

   Dim Cmd As New ADODB.Command

   Cmd.CommandTimeout = 500
   OpenStoredProcWithCommand "s_Dictation_DeleteAll", Cmd
End Sub
Public Sub DeleteHistory(DeleteOlderThanDays As Integer)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DeleteOlderThanDays)
   
   Cmd.CommandTimeout = 500
   OpenStoredProcWithCommand "s_History_DeleteByCreatedDate", Cmd
End Sub
Public Function GetCurrentTimestamp() As Date

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBDate, adParamOutput)
   
   OpenStoredProcWithCommand "s_Dictation_GetCurrentTimestamp", Cmd
   
   GetCurrentTimestamp = Cmd(0)
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
Public Function GetRoles(Roles As clsRoles, ByVal GroupId As Long, ByVal OrgId As Long) As Boolean

   Dim Rs As ADODB.Recordset
   
   Set Rs = New ADODB.Recordset
   OpenDisconnectedRecordset Rs, "Select * from GroupRoles", "GroupId=" & CStr(GroupId) & " and OrgId=" & CStr(OrgId), "", ""
   
   If Rs.EOF Then
      GetRoles = False
   Else
      Set Roles = New clsRoles
      Roles.GroupId = Rs("GroupId")
      Roles.OrgId = Rs("OrgId")
      Roles.Roles = Rs("Roles")
      Roles.Delayed = Rs("Delayed")
      Roles.DelayedHours = Rs("DelayedHours")
      GetRoles = True
   End If
   CloseRecordset Rs
End Function
Public Sub SaveRoles(Roles As clsRoles)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from GroupRoles", "GroupId=" & CStr(Roles.GroupId) & " and OrgId=" & CStr(Roles.OrgId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Rs.EOF Then
      Rs.AddNew
   End If
   Rs("GroupId") = Roles.GroupId
   Rs("OrgId") = Roles.OrgId
   Rs("Roles") = Roles.Roles
   Rs("Delayed") = SaveBool(Roles.Delayed)
   Rs("DelayedHours") = Roles.DelayedHours
   Rs.Update
   CloseRecordset Rs
End Sub

Public Sub SaveOrg(Org As clsOrg)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If Org.OrgParent = 0 Then Exit Sub     'Can't change root element
   
   SQL = BuildSQL("Select * from Org", "OrgId=" & CStr(Org.OrgId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Rs.EOF Then
      Rs.AddNew
   End If
   Rs("OrgParent") = Org.OrgParent
   Rs("OrgText") = Org.OrgText
   Rs("DictContainer") = Org.DictContainer
   Rs("ShowInTree") = Org.ShowInTree
   Rs("ShowBelow") = Org.ShowBelow
   Rs.Update
   Org.OrgId = Rs("OrgId")
   CloseRecordset Rs
End Sub
Public Sub SaveDictType(DictType As clsDictType)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from DictType", "DictTypeId=" & CStr(DictType.DictTypeId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Rs.EOF Then
      Rs.AddNew
   End If
   Rs("DictTypeId") = DictType.DictTypeId
   If Len(DictType.DictTypeText) > 0 Then
      Rs("DictTypeText") = DictType.DictTypeText
   Else
      Rs("DictTypeText") = Null
   End If
   Rs.Update
   CloseRecordset Rs
End Sub
Public Sub SaveGroup(Grp As clsGroup)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If LCase$(Grp.GroupText) = "sysadmin" Then Exit Sub   'Can't change sysadmin group
   
   SQL = BuildSQL("Select * from Groups", "GroupId=" & CStr(Grp.GroupId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Rs.EOF Then
      Rs.AddNew
   End If
   Rs("AdmOrgId") = Grp.AdmOrgId
   Rs("GroupText") = Grp.GroupText
   If Len(Grp.GroupDesc) > 0 Then
      Rs("GroupDesc") = Grp.GroupDesc
   Else
      Rs("GroupDesc") = Null
   End If
   Rs.Update
   Grp.GroupId = Rs("GroupId")
   CloseRecordset Rs
End Sub
Public Sub SaveUser(Usr As clsUser)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If LCase$(Usr.LoginName) = "sa" Then Exit Sub   'Can't change sa
   
   SQL = BuildSQL("Select * from Users", "UserId=" & CStr(Usr.UserId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Rs.EOF Then
      Rs.AddNew
   End If
   Rs("LoginName") = UCase$(Usr.LoginName)
   Rs("Password") = Usr.Password
   Rs("ShortName") = Usr.ShortName
   Rs("LongName") = Usr.LongName
   If Usr.HomeOrgId > 0 Then
      Rs("HomeOrgId") = Usr.HomeOrgId
   Else
      Rs("HomeOrgId") = Null
   End If
   Rs.Update
   Usr.UserId = Rs("UserId")
   CloseRecordset Rs
End Sub
Public Sub SaveUserHomeOrg(Usr As clsUser)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If LCase$(Usr.LoginName) = "sa" Then Exit Sub   'Can't change sa
   
   SQL = BuildSQL("Select * from Users", "UserId=" & CStr(Usr.UserId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Not Rs.EOF Then
      If Usr.HomeOrgId > 0 Then
         Rs("HomeOrgId") = Usr.HomeOrgId
      Else
         Rs("HomeOrgId") = Null
      End If
      Rs.Update
   End If
   CloseRecordset Rs
End Sub
Public Sub DeleteUser(Usr As clsUser)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If LCase$(Usr.LoginName) = "sa" Then Exit Sub   'Can't change sa
   
   SQL = BuildSQL("Select * from Users", "UserId=" & CStr(Usr.UserId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Not Rs.EOF Then
      Rs.Delete
   End If
   Rs.Update
   CloseRecordset Rs
End Sub
Public Sub DeleteOneUserGroup(UserId As Long, GroupId As Long)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from UserGroup", "UserId=" & CStr(UserId) & " and GroupId=" & CStr(GroupId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Not Rs.EOF Then
      Rs.Delete
   End If
   
   CloseRecordset Rs
End Sub

Public Sub SaveUserGroup(UserId As Long, GroupId As Long)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from UserGroup", "UserId=" & CStr(UserId) & " and GroupId=" & CStr(GroupId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   If Rs.EOF Then
      Rs.AddNew
      Rs("UserId") = UserId
      Rs("GroupId") = GroupId
      Rs.Update
   End If
   
   CloseRecordset Rs
End Sub
Function CheckInNewDemo(ByRef Dict As clsDict)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Dirty As Boolean
   
   SQL = Client.Server.BuildSQL("Dictation", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   Rs.AddNew
   
   Rs("LockedByUserId") = Null
   Rs("LockedByStation") = Null
   Rs("LockedTime") = Null
   
   Rs("NoPatient") = Dict.NoPatient
   Rs("PatId") = Dict.Pat.PatId
   Rs("PatName") = Dict.Pat.PatName
   Rs("Created") = Dict.Created
   If Dict.Changed > 0 Then
      Rs("Changed") = Dict.Changed
   End If
   Rs("OrgId") = Dict.OrgId
   Rs("DictTypeId") = Dict.DictTypeId
   Rs("PriorityId") = Dict.PriorityId
   Rs("ExpiryDate") = Dict.ExpiryDate
   If Len(Dict.Txt) > 0 Then
      Rs("Txt") = Dict.Txt
   End If
   
   If Len(Dict.LocalFilename) > 0 Then
      Rs("SoundStored") = True
      Rs("SoundCreated") = Dict.Created
      Rs("SoundLength") = Dict.SoundLength
   End If
               
   Rs("StatusId") = Dict.StatusId
   Rs("AuthorId") = Dict.AuthorId
   If Dict.TranscriberId > 0 Then
      Rs("TranscriberId") = Dict.TranscriberId
      Rs("TranscribedDate") = Dict.TranscribedDate
      If Dict.TranscriberOrgId > 0 Then
         Rs("TranscriberOrgId") = Dict.TranscriberOrgId
      End If
   Else
      Rs("TranscriberId") = Null
   End If
   Rs("TimeStamp") = Now
   Rs.Update
   Dict.DictId = Rs("DictId")
   If Len(Dict.LocalFilename) > 0 Then
      CopyFileToSoundStorage Dict.LocalFilename, Dict.DictId
      KillFileIgnoreError Dict.LocalFilename
      Dict.LocalFilename = ""
   End If
   CheckInNewDemo = 0
   
   CloseRecordset Rs
End Function
Public Sub CreateDictIdList(OrgId As Long, AuthorId As Long, TranscriberId As Long, Optional PatId As String = "", Optional PatId2 As String = "")

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
Public Function DictIdGetNext(ByRef DictId As Long) As Boolean

   If RsDictIdList.EOF Then
      DictIdGetNext = False
      RsDictIdList.Close
      Set RsDictIdList = Nothing
   Else
      DictId = RsDictIdList("DictId")
      RsDictIdList.MoveNext
      DictIdGetNext = True
   End If
End Function
Public Sub CreateHist(OrgId As Long, HistYear As Integer, HistType As HistTypeEnum)

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , OrgId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , HistYear)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , HistType)
   
   OpenStoredProcWithRsResultAndCommand RsHist, "s_History_CreateList", Cmd
End Sub
Public Function HistGetNext(ByRef Hist As clsHistory) As Boolean

   If RsHist.EOF Then
      HistGetNext = False
      RsHist.Close
      Set RsHist = Nothing
   Else
      ConvertHistRsToClass RsHist, Hist
      RsHist.MoveNext
      HistGetNext = True
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
Public Sub CreateStat(OrgId As Long, ByRef Stat As clsStat, StatusIdStart As Long, StatusIdEnd As Long, _
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
Public Function GetSoundfilePath(DictId As Long) As String

   GetSoundfilePath = DictationStorageSoundPath & CStr(DictId Mod 10) & "\" & CStr(DictId)
End Function
Public Function NeedLoginInfo() As Boolean

   NeedLoginInfo = mAuthMethod < 2
End Function

Public Function UserLogin(User As clsUser, _
                          LoginName As String, _
                          Password As String, _
                          NewPassword As String, _
                          ExtSystem As String, _
                          ExtPassword As String) As Integer

   Dim Rs As New ADODB.Recordset
   Dim SQL As String
   Dim RightPassword As String
    
   If mAuthMethod = 2 Then
      If Not SetUserFromSQLLoginName(User) Then
         UserLogin = 100
         Exit Function
      End If
      UserLogin = 0
      Exit Function
   End If
   
   If mAuthMethod < 2 Then
      SQL = BuildSQL("SELECT * FROM Users", "LoginName='" & UCase$(LoginName) & "'", "", "")
      OpenRecordset Rs, SQL
      If Not Rs.EOF Then
         RightPassword = Rs("Password")
      Else
         RightPassword = ""
      End If
      
      If Not Client.SysSettings.LoginPasswordCaseSensitive Then
         RightPassword = UCase$(RightPassword)
         Password = UCase$(Password)
      End If
      
      If Not Rs.EOF Then
         If nvl(Rs("UserLockedUntil"), 0) > Now Then
            UserLogin = 1
         Else
            If (Len(ExtPassword) > 0 And ExtPassword = Client.ExtSystemMgr.GetExtSystemPassword(ExtSystem)) Or RightPassword = Password Then
               User.UserId = Rs("UserId")
               User.LoginName = Rs("LoginName")
               User.ShortName = Rs("ShortName")
               User.LongName = Rs("LongName")
               User.HomeOrgId = nvl(Rs("HomeOrgid"), 0)
               User.Password = nvl(Rs("Password"), "")
               Rs("LastSuccessfulLogin") = Now
               Rs("LoginFailureCount") = 0
               Rs("UserLockedUntil") = Null
               If Len(NewPassword) > 0 Then
                  Rs("Password") = NewPassword
               End If
               UserLogin = 0
            Else
               If Rs("LoginFailureCount") >= Client.SysSettings.LoginMaxRetries Then
                  Rs("LoginFailureCount") = 0
                  Rs("UserLockedUntil") = DateAdd("n", Client.SysSettings.LoginLockoutMinutes, Now)
                  UserLogin = 1
               Else
                  Rs("LoginFailureCount") = Rs("LoginFailureCount") + 1
                  UserLogin = 2
               End If
            End If
         End If
         Rs.Update
      Else
         UserLogin = 3
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
Public Sub CreateLoggList(StartTime As Date, EndTime As Date, StartLevel As Integer, EndLevel As Integer)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(StartTime))
   Cmd.Parameters.Append Cmd.CreateParameter(, adDBTimeStamp, adParamInput, , SaveDate(EndTime))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StartLevel)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , EndLevel)
   OpenStoredProcWithRsResultAndCommand RsLogg, "s_Logg_GetByDateAndLevel", Cmd
End Sub
Public Function LoggListGetNext(ByRef Logg As clsLogg) As Boolean

   If RsLogg.EOF Then
      RsLogg.Close
      Set RsLogg = Nothing
      LoggListGetNext = False
      Exit Function
   Else
      ConvertLoggRsToClass RsLogg, Logg
      RsLogg.MoveNext
      LoggListGetNext = True
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
Public Sub LoggInsert(Logg As clsLogg)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Logg.LoggId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Logg.LoggLevel)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Logg.UserId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Logg.StationId, 50))
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Logg.DictId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Logg.LoggData, 50))
   
   OpenStoredProcWithCommand "s_Logg_Insert", Cmd
End Sub
Public Sub CreateDictAuditList(DictId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
   OpenStoredProcWithRsResultAndCommand RsDictAudit, "s_DictAudit_GetByDictId", Cmd
End Sub
Public Function DictAuditGetNext(ByRef Audit As clsDictAudit) As Boolean

   If RsDictAudit.EOF Then
      RsDictAudit.Close
      Set RsDictAudit = Nothing
      DictAuditGetNext = False
      Exit Function
   Else
      ConvertDictAuditRsToClass RsDictAudit, Audit
      RsDictAudit.MoveNext
      DictAuditGetNext = True
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
Public Sub CreateExtSystemList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystem, "s_ExtSystem_Get"
End Sub
Public Function ExtSystemGetNext(ByRef E As clsExtSystem) As Boolean

   If RsExtSystem.EOF Then
      RsExtSystem.Close
      Set RsExtSystem = Nothing
      ExtSystemGetNext = False
      Exit Function
   Else
      ConvertExtSystemRsToClass RsExtSystem, E
      RsExtSystem.MoveNext
      ExtSystemGetNext = True
      Exit Function
   End If
End Function
Private Sub ConvertExtSystemRsToClass(ByRef Rs As ADODB.Recordset, ByRef E As clsExtSystem)

   Dim S As String
   
   Set E = New clsExtSystem
   With E
      .ExtSystem = Rs("ExtSystem")
      On Error Resume Next    'Ok if field is not in table
      S = nvl(Rs("ExtSystemPw"), "")
      If Len(S) > 0 Then
         .ExtSystemPw = RC4(S, ct_ext_crypto_key)
      Else
         .ExtSystemPw = ""
      End If
      On Error GoTo 0
   End With
End Sub
Public Sub CreateExtSystemDictTypeList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystemDictType, "s_ExtSystemDictType_Get"
End Sub
Public Function ExtSystemDictTypeGetNext(ByRef E As clsExtSystemDictType) As Boolean

   If RsExtSystemDictType.EOF Then
      RsExtSystemDictType.Close
      Set RsExtSystemDictType = Nothing
      ExtSystemDictTypeGetNext = False
      Exit Function
   Else
      ConvertExtSystemDictTypeRsToClass RsExtSystemDictType, E
      RsExtSystemDictType.MoveNext
      ExtSystemDictTypeGetNext = True
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
Public Sub CreateExtSystemOrgList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystemOrg, "s_ExtSystemOrg_Get"
End Sub
Public Function ExtSystemOrgGetNext(ByRef E As clsExtSystemOrg) As Boolean

   If RsExtSystemOrg.EOF Then
      RsExtSystemOrg.Close
      Set RsExtSystemOrg = Nothing
      ExtSystemOrgGetNext = False
      Exit Function
   Else
      ConvertExtSystemOrgRsToClass RsExtSystemOrg, E
      RsExtSystemOrg.MoveNext
      ExtSystemOrgGetNext = True
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
Public Sub CreateExtSystemPriorityList()

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   OpenStoredProcWithRsResult RsExtSystemPriority, "s_ExtSystemPriority_Get"
End Sub
Public Function ExtSystemPriorityGetNext(ByRef E As clsExtSystemPriority) As Boolean

   If RsExtSystemPriority.EOF Then
      RsExtSystemPriority.Close
      Set RsExtSystemPriority = Nothing
      ExtSystemPriorityGetNext = False
      Exit Function
   Else
      ConvertExtSystemPriorityRsToClass RsExtSystemPriority, E
      RsExtSystemPriority.MoveNext
      ExtSystemPriorityGetNext = True
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
Public Sub DictAuditInsert(Audit As clsDictAudit)

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Audit.DictId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Audit.AuditType)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Audit.DictStatus)
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , SaveNullIfZero(Audit.UserId))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, SaveNullZeroLength(Audit.StationId, 50))
   
   OpenStoredProcWithCommand "s_DictAudit_Insert", Cmd
End Sub
Public Sub CreateOrgList()

   OpenDisconnectedRecordset RsOrgList, "SELECT * FROM Org", "", "", ""
   Set RsOrgListClone = RsOrgList.Clone
   OpenDisconnectedRecordset RsUserRole, "SELECT * FROM v_UserRole", "Userid=" & Client.User.UserId, "", ""
   
End Sub
Public Function OrgListGetNext(ByRef Org As clsOrg) As Boolean

   Do While True
      If RsOrgList.EOF Then
         RsOrgList.Close
         Set RsOrgList = Nothing
         RsOrgListClone.Close
         Set RsOrgListClone = Nothing
         RsUserRole.Close
         Set RsUserRole = Nothing
         OrgListGetNext = False
         Exit Function
      Else
         ConvertOrgRsToClass RsOrgList, Org
         RsOrgList.MoveNext
         If Len(Org.Roles.Roles) > 0 Then
            OrgListGetNext = True
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
      If Len(RolesFromParent.Roles) = 0 Then
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

Public Sub CreateDictList(ByVal OrgId As Long, ByVal AuthorId As Long, ByVal TranscriberId As Long, StatusIdStart As Long, StatusIdEnd As Long, TimeStamp As Double, Optional PatId As String = "", Optional PatId2 As String = "")

   Dim Cmd As New ADODB.Command
   
   DictListCount = 0
   
   If Len(PatId) > 0 Or Len(PatId2) > 0 Then
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 12, Left$(PatId, 12))
      Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 12, Left$(PatId2, 12))
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdStart)
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , StatusIdEnd)
      
      OpenStoredProcWithRsResultAndCommand RsDictList, "s_Dictation_GetListByPatid", Cmd
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
Public Sub CreateDictListFromSQL(SQL As String)

   DictListCount = 0
   OpenRecordset RsDictList, SQL
End Sub

Public Function DictListGetNext(ByRef Dict As clsDict) As Boolean

   If RsDictList.EOF Then
      Debug.Print "DictListCount: " & CStr(DictListCount)
      DictListGetNext = False
      RsDictList.Close
      Set RsDictList = Nothing
   Else
      DictListCount = DictListCount + 1
      ConvertDictRsToClass RsDictList, Dict
      RsDictList.MoveNext
      DictListGetNext = True
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
Public Sub CreateDictTypeList()

   OpenStoredProcWithRsResult RsDictTypeList, "s_DictType_Get"
End Sub
Public Function DictTypeListGetNext(DictType As clsDictType) As Boolean

   If RsDictTypeList.EOF Then
      DictTypeListGetNext = False
      RsDictTypeList.Close
      Set RsDictTypeList = Nothing
   Else
      ConvertDictTypeRsToClass RsDictTypeList, DictType
      RsDictTypeList.MoveNext
      DictTypeListGetNext = True
   End If
End Function
Private Sub ConvertDictTypeRsToClass(ByRef Rs As ADODB.Recordset, ByRef DictType As clsDictType)

   Set DictType = New clsDictType

   DictType.DictTypeId = Rs("DictTypeId")
   DictType.DictTypeText = Rs("DictTypeText")
End Sub
Public Sub CreateOrgDictTypeList()

   OpenStoredProcWithRsResult RsOrgDictTypeList, "s_OrgDictType_Get"
End Sub
Public Function OrgDictTypeListGetNext(OrgDictType As clsOrgDictType) As Boolean

   If RsOrgDictTypeList.EOF Then
      OrgDictTypeListGetNext = False
      RsOrgDictTypeList.Close
      Set RsOrgDictTypeList = Nothing
   Else
      ConvertOrgDictTypeRsToClass RsOrgDictTypeList, OrgDictType
      RsOrgDictTypeList.MoveNext
      OrgDictTypeListGetNext = True
   End If
End Function
Private Sub ConvertOrgDictTypeRsToClass(ByRef Rs As ADODB.Recordset, ByRef OrgDictType As clsOrgDictType)

   Set OrgDictType = New clsOrgDictType

   OrgDictType.OrgId = Rs("OrgId")
   OrgDictType.DictTypeId = Rs("DictTypeId")
   OrgDictType.Def = Rs("Def") <> 0
End Sub
Public Sub DeleteOrgDictTypeByOrgId(OrgId As Long)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from OrgDictType", "OrgId=" & CStr(OrgId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   Do While Not Rs.EOF
      Rs.Delete
      Rs.MoveNext
   Loop
   
   CloseRecordset Rs
End Sub
Public Sub SaveOrgDictType(OrgId As Long, DictTypeId As Integer, DictTypeDefault As Boolean)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from OrgDictType", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   Rs.AddNew
   Rs("OrgId") = OrgId
   Rs("DictTypeId") = DictTypeId
   Rs("Def") = SaveBool(DictTypeDefault)
   Rs.Update
   
   CloseRecordset Rs
End Sub
Public Sub CreateOrgPriorityList()

   OpenStoredProcWithRsResult RsOrgPriorityList, "s_OrgPriority_Get"
End Sub
Public Function OrgPriorityListGetNext(OrgPriority As clsOrgPriority) As Boolean

   If RsOrgPriorityList.EOF Then
      OrgPriorityListGetNext = False
      RsOrgPriorityList.Close
      Set RsOrgPriorityList = Nothing
   Else
      ConvertOrgPriorityRsToClass RsOrgPriorityList, OrgPriority
      RsOrgPriorityList.MoveNext
      OrgPriorityListGetNext = True
   End If
End Function
Private Sub ConvertOrgPriorityRsToClass(ByRef Rs As ADODB.Recordset, ByRef OrgPriority As clsOrgPriority)

   Set OrgPriority = New clsOrgPriority

   OrgPriority.OrgId = Rs("OrgId")
   OrgPriority.PriorityId = Rs("PriorityId")
   OrgPriority.Def = Rs("Def") <> 0
End Sub
Public Sub DeleteOrgPriorityByOrgId(OrgId As Long)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from OrgPriority", "OrgId=" & CStr(OrgId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   Do While Not Rs.EOF
      Rs.Delete
      Rs.MoveNext
   Loop
   
   CloseRecordset Rs
End Sub
Public Sub SaveOrgPriority(OrgId As Long, PriorityId As Integer, PriorityDefault As Boolean)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from OrgPriority", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   Rs.AddNew
   Rs("OrgId") = OrgId
   Rs("PriorityId") = PriorityId
   Rs("Def") = SaveBool(PriorityDefault)
   Rs.Update
   
   CloseRecordset Rs
End Sub
Public Sub CreateGroupList()

   OpenStoredProcWithRsResult RsGroupList, "s_Group_Get"
End Sub
Public Function GroupListGetNext(Grp As clsGroup) As Boolean

   If RsGroupList.EOF Then
      GroupListGetNext = False
      RsGroupList.Close
      Set RsGroupList = Nothing
   Else
      ConvertGroupRsToClass RsGroupList, Grp
      RsGroupList.MoveNext
      GroupListGetNext = True
   End If
End Function
Private Sub ConvertGroupRsToClass(ByRef Rs As ADODB.Recordset, ByRef Grp As clsGroup)

   Set Grp = New clsGroup

   Grp.GroupId = Rs("GroupId")
   Grp.GroupText = Rs("GroupText")
   Grp.GroupDesc = nvl(Rs("GroupDesc"), "")
   Grp.AdmOrgId = nvl(Rs("AdmOrgId"), 0)
End Sub
Public Sub CreateGroupListForUser(UserId As Long)

   OpenDisconnectedRecordset RsGroupListForUser, "SELECT * FROM UserGroup", "Userid=" & CStr(UserId), "", ""
End Sub
Public Function GroupListForUserGetNext(ByRef GroupId As Long) As Boolean

   If RsGroupListForUser.EOF Then
      GroupListForUserGetNext = False
      RsGroupListForUser.Close
      Set RsGroupListForUser = Nothing
   Else
      GroupId = RsGroupListForUser("GroupId")
      RsGroupListForUser.MoveNext
      GroupListForUserGetNext = True
   End If
End Function
Public Sub CreateUserList()

   OpenDisconnectedRecordset RsUsers, "SELECT * FROM Users", "", "", "ShortName ASC"
End Sub
Public Function UserListGetNext(Usr As clsUser) As Boolean

   If RsUsers.EOF Then
      UserListGetNext = False
      CloseRecordset RsUsers
   Else
      ConvertUsersRsToClass RsUsers, Usr
      RsUsers.MoveNext
      UserListGetNext = True
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
End Sub
Public Sub CreatePriorityList()

   OpenDisconnectedRecordset RsPriorityList, "SELECT * FROM Priority", "", "", "Days ASC"
End Sub
Public Function PriorityListGetNext(Priority As clsPriority) As Boolean

   If RsPriorityList.EOF Then
      PriorityListGetNext = False
      RsPriorityList.Close
      Set RsPriorityList = Nothing
   Else
      ConvertPriorityRsToClass RsPriorityList, Priority
      RsPriorityList.MoveNext
      PriorityListGetNext = True
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

Private Sub Class_Initialize()

   Set Settings = New clsSettings
   
   Settings.IniFilePath = App.Path & "\" & CommandValue("Profile", "CareTalk") & ".ini"
End Sub

Private Sub Class_Terminate()

   DictationStorageClose
   
   Set Settings = Nothing
   
End Sub

Public Function DictationStorageOpen(Server As String, _
                                     Database As String, _
                                     UserId As String, _
                                     Password As String) As Boolean

   Dim Rs As New ADODB.Recordset
   Dim CS As String
   Dim Provider As String
   Dim Cred As String
   Dim CurrUserid As String
   Dim CurrPassword As String
   Dim UseTrustedInSettings As Boolean

   SQLDateStringFormat = Settings.GetString("DictationStorage", "SQLDateStringFormat", "mm/dd/yyyy")
   SQLDateDelimiter = Settings.GetString("DictationStorage", "SQLDateDelimiter", "'")

   DictationStoragePath = Settings.GetFolder("DictationStorage", "Folder", "")
   DictationStorageSoundPath = Settings.GetFolder("DictationStorage", "SoundFolder", DictationStoragePath)
   
   Provider = Settings.GetString("DictationStorage", "Provider", "sqloledb")
   If Len(Server) = 0 Then
      SQLServer = Settings.GetString("DictationStorage", "Server", "")
   Else
      SQLServer = Server
   End If
      
   If Len(Database) = 0 Then
      SQLDatabase = Settings.GetString("DictationStorage", "Database", "CareTalk")
   Else
      SQLDatabase = Database
   End If
   
   UseTrustedInSettings = Settings.GetBool("DictationStorage", "TrustedConnection", False)
   
   CS = "Provider=" & Provider
   CS = CS & ";Server=" & SQLServer
   CS = CS & ";Database=" & SQLDatabase
   
   CS = Settings.GetString("DictationStorage", "ConnectionString", CS)
     
   If UseTrustedInSettings Then
      Cred = "Trusted_Connection=yes"
      mConnectionString = CS & ";" & Cred
      mAuthMethod = 2
   Else
      Cred = "User Id=" & DBUserid & ";Password=" & DBPassword
      mConnectionString = CS & ";" & Cred
      mAuthMethod = 0
   End If
   Db_ConnectionOpen
   DictationStorageOpen = mStorageOpened
End Function
Private Sub Db_ConnectionOpen()

   On Error Resume Next
   Db_ConnectionCloseIfOpened
   
   Set Db = New ADODB.Connection
   Db.ConnectionString = mConnectionString
   Db.CommandTimeout = 180
   Db.Open
   mStorageOpened = Err = 0

   If mStorageOpened Then
      Db_LoginToAppRole
   End If
End Sub
Private Sub Db_PrepareConnection()

   If Db Is Nothing Then
      Db_ConnectionOpen
   ElseIf Db.State = adStateClosed Then
      Db_ConnectionOpen
   End If
End Sub
Private Sub Db_TryToRepaireConnection()

   Dim I As Integer
   
   I = 3
   Do
      MsgWaitObj 2000
      Db_ConnectionOpen
      I = I - 1
   Loop Until I <= 0 Or mStorageOpened
End Sub
Private Sub Db_ConnectionCloseIfOpened()

   On Error Resume Next
   If Not Db Is Nothing Then
      Db.Close
   End If
   Set Db = Nothing
End Sub
Private Sub Db_LoginToAppRole()

   If mStorageOpened Then
      If Not Db_GetAppInfo() Then
         DictationStorageClose
      End If
      If Not Db_ActivateAppRole() Then
         DictationStorageClose
      End If
   End If
End Sub

Private Function Db_ActivateAppRole() As Boolean

   Dim App_password As String
   
   If Len(mCryptoAppRolePassword) = 0 Then
      App_password = ct_app_default_password
   Else
      App_password = RC4(mCryptoAppRolePassword, ct_app_crypto_key)
   End If

   On Error GoTo ActivateAppRole_Err
   Db.Execute "EXEC sp_setapprole 'ct_app', '" & App_password & "'"
   Db_ActivateAppRole = True
   Exit Function
   
ActivateAppRole_Err:
   Db_ActivateAppRole = False
   Exit Function
End Function
Private Function Db_GetAppInfo() As Boolean

   Dim Rs As ADODB.Recordset

   On Error GoTo GetAppInfo_Err
   OpenStoredProcWithRsResult Rs, "s_AppInfo"
   If Not Rs.EOF Then
      mCryptoAppRolePassword = nvl(Rs("AppPw"), "")
   End If
   CloseRecordset Rs
   Db_GetAppInfo = True
   Exit Function
   
GetAppInfo_Err:
   Db_GetAppInfo = False
   CloseRecordset Rs
   Exit Function
End Function
Private Function RC4(ByVal Expression As String, ByVal Password As String) As String
   On Error Resume Next
   Dim RB(0 To 255) As Integer, x As Long, y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
   If Len(Password) = 0 Then
       Exit Function
   End If
   If Len(Expression) = 0 Then
       Exit Function
   End If
   If Len(Password) > 256 Then
       Key() = StrConv(Left$(Password, 256), vbFromUnicode)
   Else
       Key() = StrConv(Password, vbFromUnicode)
   End If
   For x = 0 To 255
       RB(x) = x
   Next x
   x = 0
   y = 0
   Z = 0
   For x = 0 To 255
       y = (y + RB(x) + Key(x Mod Len(Password))) Mod 256
       Temp = RB(x)
       RB(x) = RB(y)
       RB(y) = Temp
   Next x
   x = 0
   y = 0
   Z = 0
   ByteArray() = StrConv(Expression, vbFromUnicode)
   For x = 0 To Len(Expression)
       y = (y + 1) Mod 256
       Z = (Z + RB(y)) Mod 256
       Temp = RB(y)
       RB(y) = RB(Z)
       RB(Z) = Temp
       ByteArray(x) = ByteArray(x) Xor (RB((RB(y) + RB(Z)) Mod 256))
   Next x
   RC4 = StrConv(ByteArray, vbUnicode)
End Function

Public Sub DictationStorageClose()

   On Error Resume Next
   mStorageOpened = False
   Db_ConnectionCloseIfOpened
   DictationStorageSoundPath = ""
End Sub

Public Function CheckOutDict(ByRef Dict As clsDict, ByVal DictId As Long, ByVal WithSound As Boolean, ByVal LocalTempFileName As String) As Integer

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   RaiseEvent UIStatusSet(Client.Texts.Txt(1260101, "Diktatet hämtas"), True)
   If LockDictation(DictId) Then
      
      Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)
        
      OpenStoredProcWithRsResultAndCommand Rs, "s_Dictation_GetById", Cmd
   
      If Rs.EOF Then
         CloseRecordset Rs
         CheckOutDict = 2
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
      
      If WithSound And Len(LocalTempFileName) > 0 And Dict.StatusId < SoundDeleted Then
         Dict.LocalFilename = LocalTempFileName
         CopyFileFromSoundStorage DictId, LocalTempFileName
      End If
     
      CheckOutDict = 0
   Else
      CheckOutDict = 1
   End If
CheckOutDict_Exit:
   RaiseEvent UIStatusClear
End Function
Public Function GetDictIdFromExtDictId(ExtSystem As String, ExtDictId As String) As Long

   If Len(ExtSystem) = 0 Or Len(ExtDictId) = 0 Then Exit Function
   
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamReturnValue)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, SaveNullZeroLength(ExtSystem, 10))
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, SaveNullZeroLength(ExtDictId, 20))

   OpenStoredProcWithCommand "s_Dictation_GetDictIdFromExtDictId", Cmd
   
   GetDictIdFromExtDictId = Cmd(0)
   
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
Public Function CheckInDict(ByRef Dict As clsDict, DiscardChanges As Boolean)

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   Dim Dirty As Boolean
   
   Client.Trace.AddRow Trace_Level_Full, "CID", "CIDA", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   RaiseEvent UIStatusSet(Client.Texts.Txt(1260102, "Diktatet sparas"), True)
      
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Dict.DictId)
   
   Client.Trace.AddRow Trace_Level_Full, "CID", "CIDB", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   OpenStoredProcWithRsResultAndCommand Rs, "s_Dictation_GetById", Cmd
   Client.Trace.AddRow Trace_Level_Full, "CID", "CIDC", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   
   If Rs.EOF Then
      Client.Trace.AddRow Trace_Level_Full, "CID", "CIDD", "", CStr(Dict.DictId), CStr(Dict.StatusId)
      CheckInDict = 1
   Else
      Client.Trace.AddRow Trace_Level_Full, "CID", "CIDE", "", CStr(Dict.DictId), CStr(Dict.StatusId)
      If Rs("LockedByUserId") <> Client.User.UserId Then
         Client.Trace.AddRow Trace_Level_Full, "CID", "CIDF", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         CheckInDict = 2
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
               
               If Len(Dict.LocalFilename) > 0 Then
                  If Dict.SoundDirty Then
                     CopyFileToSoundStorage Dict.LocalFilename, Dict.DictId
                  End If
               End If
               
               If Dirty Or Dict.SoundDirty Then
                  Dict.Changed = Now
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
            If Rs("StatusId") < WaitForSign And Dict.StatusId >= WaitForSign Then
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
            CheckInDict = CheckInDictInternal(Dict)
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDM", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         Else
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDN", "", CStr(Dict.DictId), CStr(Dict.StatusId)
            UnLockDictationWithNewStatus Dict.DictId, Dict.StatusId
            Client.Trace.AddRow Trace_Level_Full, "CID", "CIDO", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         End If
      End If
   End If
   CloseRecordset Rs
   RaiseEvent UIStatusClear
End Function
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
   OpenStoredProcWithCommand "s_Dictation_CheckIn2", Cmd
   Client.Trace.AddRow Trace_Level_Full, "CIDI", "CIDIC", "", CStr(Dict.DictId), CStr(Dict.StatusId)
   Dict.DictId = Cmd(0)
   Client.Trace.AddRow Trace_Level_Full, "CIDI", "CIDID", "", CStr(Dict.DictId), CStr(Dict.StatusId)
End Function
Public Function CheckInNewDict(ByRef Dict As clsDict) As Integer

   RaiseEvent UIStatusSet(Client.Texts.Txt(1260102, "Diktatet sparas"), True)
      Dict.AuthorId = Client.User.UserId
      Dict.AuthorPos = Dict.CurrentPos
      CheckInDictInternal Dict
      
      If Len(Dict.LocalFilename) > 0 Then
         CopyFileToSoundStorage Dict.LocalFilename, Dict.DictId
         KillFileIgnoreError Dict.LocalFilename
         Dict.LocalFilename = ""
      End If
      CheckInNewDict = 0
   RaiseEvent UIStatusClear
End Function
Public Sub UnlockDict(DictId As Long)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , DictId)

   OpenStoredProcWithCommand "s_Dictation_Unlock", Cmd
   
End Sub

Private Sub AddRole(SumRoles As clsRoles, ByVal Roles2 As clsRoles)

   Dim I As Integer
   Dim Letter As String
   Dim R2 As String
   Dim DoWeAlreadyHaveARole As Boolean
   
   SumRoles.GroupId = Roles2.GroupId
   SumRoles.OrgId = Roles2.OrgId
   
   DoWeAlreadyHaveARole = Len(SumRoles.Roles) > 0
   
   R2 = UCase$(Roles2.Roles)
   For I = 1 To Len(R2)
      Letter = mId$(R2, I, 1)
      If InStr(SumRoles.Roles, Letter) = 0 Then
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
Private Sub OpenDisconnectedRecordset(ByRef Rs As ADODB.Recordset, ByVal SelectFrom As String, ByVal WhereClause As String, ByVal GroupByClause As String, ByVal OrderClause As String)

   Dim T As Double
   Dim SQL As String
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
  
OpenDisconnectedRecordset_Start:

   On Error GoTo OpenDisconnectedRecordset_err
   Db_PrepareConnection
   
   CloseRecordset Rs
   Set Rs = New ADODB.Recordset
   Set Rs.ActiveConnection = Db
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   SQL = BuildSQL(SelectFrom, WhereClause, GroupByClause, OrderClause)
   T = Timer
   Rs.Open SQL
   Debug.Print SQL, Format$(Timer - T, "0.00")
   Set Rs.ActiveConnection = Nothing
   Exit Sub
   
OpenDisconnectedRecordset_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume OpenDisconnectedRecordset_Start
   Else
      Error ErrVal
   End If
End Sub
Private Sub OpenRecordset(ByRef Rs As ADODB.Recordset, ByRef SQL As String)

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   
OpenRecordset_Start:

   On Error GoTo OpenRecordset_err
   Db_PrepareConnection
   
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
   End If
   
   T = Timer
   Rs.Open SQL, Db, adOpenDynamic, adLockOptimistic
   Debug.Print SQL, Format$(Timer - T, "0.00")
   Exit Sub
   
OpenRecordset_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume OpenRecordset_Start
   Else
      Error ErrVal
   End If

End Sub
Public Sub CloseRecordset(Rs As ADODB.Recordset)

   On Error Resume Next
   Rs.Close
   Set Rs = Nothing
End Sub

Private Sub OpenStoredProc(SpName As String)

   Dim Cmd As ADODB.Command
   
   Set Cmd = New ADODB.Command
   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   ExecuteCommand Cmd
End Sub
Private Sub OpenStoredProcWithRsResult(ByRef Rs As ADODB.Recordset, SpName As String)

   Dim Cmd As ADODB.Command
   
   Set Cmd = New ADODB.Command
   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   ExecuteCommandWithRsResult Rs, Cmd
End Sub
Private Sub OpenStoredProcWithRsResultAndCommand(ByRef Rs As ADODB.Recordset, SpName As String, ByRef Cmd As ADODB.Command)

   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   ExecuteCommandWithRsResult Rs, Cmd
End Sub
Private Sub OpenStoredProcWithCommand(SpName As String, ByRef Cmd As ADODB.Command)

   Cmd.CommandText = SpName
   Cmd.CommandType = adCmdStoredProc
   
   ExecuteCommand Cmd
End Sub
Private Sub ExecuteCommand(ByRef Cmd As ADODB.Command)

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   
ExecuteCommand_Start:

   On Error GoTo ExecuteCommand_err
   Db_PrepareConnection
   
   Cmd.ActiveConnection = Db

   T = Timer
   Cmd.Execute
   Debug.Print Cmd.CommandText, Format$(Timer - T, "0.00")
   Exit Sub
   
ExecuteCommand_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume ExecuteCommand_Start
   Else
      Error ErrVal
   End If
End Sub
Private Sub ExecuteCommandWithRsResult(ByRef Rs As ADODB.Recordset, ByRef Cmd As ADODB.Command)

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   
ExecuteCommandWithRsResult_Start:

   On Error GoTo ExecuteCommandWithRsResult_err
   Db_PrepareConnection
   
   Cmd.ActiveConnection = Db
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   'Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockOptimistic

   Debug.Print Cmd.ActiveConnection.State
   'If Cmd.ActiveConnection.State <> adStateOpen Then Stop '!!!
   T = Timer
   Set Rs = Cmd.Execute
   Debug.Print Cmd.CommandText, Format$(Timer - T, "0.00")
   Exit Sub
   
ExecuteCommandWithRsResult_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume ExecuteCommandWithRsResult_Start
   Else
      Error ErrVal
   End If
End Sub

Sub CheckSQLInjection(ByVal SQLstatement As String)

   'Err.Raise 30000   !!!
End Sub
Public Function BuildSQL(SelectFrom As String, WhereClause As String, GroupByClause As String, OrderClause As String)

   Dim Res As String
   
   CheckSQLInjection WhereClause
   CheckSQLInjection GroupByClause
   CheckSQLInjection OrderClause
   
   Res = SelectFrom
   If Len(WhereClause) > 0 Then
      Res = Res & " WHERE " & WhereClause
   End If
   If Len(GroupByClause) > 0 Then
      Res = Res & " GROUP BY " & GroupByClause
   End If
   If Len(OrderClause) > 0 Then
      Res = Res & " ORDER BY " & OrderClause
   End If
   BuildSQL = Res
End Function
Public Sub AddString(Dest As String, AddPhrase As String, Addendum As String)

   If Len(Addendum) = 0 Then Exit Sub
   If Len(Dest) = 0 Then
      Dest = Addendum
   Else
      Dest = Dest & " " & AddPhrase & " " & Addendum
   End If
End Sub
Public Sub WriteUserData(Ty As String, Ke As String, Va As String)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Client.User.UserId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, Ty)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, Ke)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 4000, Va)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, ApplicationVersion)
   
   OpenStoredProcWithCommand "s_UserData_Set", Cmd

End Sub
Public Function ReadUserData(Ty As String, Ke As String, Def As String, ByRef version) As String

   Dim Rs As ADODB.Recordset
   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adInteger, adParamInput, , Client.User.UserId)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 10, Ty)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 20, Ke)
   
   OpenStoredProcWithRsResultAndCommand Rs, "s_UserData_Get", Cmd
   
   If Not Rs.EOF Then
      ReadUserData = Rs("Va")
      version = Rs("Version")
   End If
   CloseRecordset Rs
End Function
Public Sub CreateSysSettingsList(ByVal Ty As String)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Ty)
   
   OpenStoredProcWithRsResultAndCommand RsSysSettingsList, "s_SysSettings_GetByTy", Cmd
End Sub
Public Function SysSettingsListGetNext(ByRef Se As String, ByRef Ke As String, ByRef Va As String) As Boolean

   If RsSysSettingsList.EOF Then
      SysSettingsListGetNext = False
      RsSysSettingsList.Close
      Set RsSysSettingsList = Nothing
   Else
      Se = RsSysSettingsList("Se")
      Ke = RsSysSettingsList("Ke")
      Va = RsSysSettingsList("Va")
      RsSysSettingsList.MoveNext
      SysSettingsListGetNext = True
   End If
End Function
Public Function WriteSysSettings(Ty As String, Se As String, Ke As String, S As String)

   Dim Cmd As New ADODB.Command
   
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Ty)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Se)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 50, Ke)
   Cmd.Parameters.Append Cmd.CreateParameter(, adVarChar, adParamInput, 200, S)
   
   OpenStoredProcWithCommand "s_SysSettings_Set", Cmd
   
End Function
Public Sub WriteStationData(Ty As String, Ke As String, Va As String)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from StationData", "Ty='" & Ty & "' and Ke='" & Ke & "' and StationId='" & Client.Station.Id & "'", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   If Rs.EOF Then
      Rs.AddNew
      Rs("Stationid") = Client.Station.Id
      Rs("Ty") = Ty
      Rs("Ke") = Ke
   End If
   Rs("Va") = Va
   Rs("Version") = ApplicationVersion
   Rs.Update
   CloseRecordset Rs
End Sub
Public Function ReadStationData(Ty As String, Ke As String, Def As String, ByRef version) As String

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from StationData", "Ty='" & Ty & "' and Ke='" & Ke & "' and StationId='" & Client.Station.Id & "'", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   OpenRecordset Rs, SQL
   If Rs.EOF Then
      ReadStationData = Def
      version = ""
   Else
      ReadStationData = Rs("Va")
      version = Rs("Version")
   End If
   CloseRecordset Rs
End Function

Private Sub CopyFileToSoundStorage(SourcePath As String, DictId As Long)

   Dim SoundPath As String

   RaiseEvent UIStatusSetSub(Client.Texts.Txt(1260103, "Ljudet sparas..."))
   If Client.SysSettings.SoundUseImageField Then
      CopyFileToSoundTable SourcePath, DictId
   Else
      SoundPath = GetSoundfilePath(DictId)
      On Error Resume Next
      SetAttr SoundPath, vbNormal
      On Error GoTo 0
      
      FileCopyWithRetry SourcePath, SoundPath
      
      On Error Resume Next
      SetAttr SoundPath, vbReadOnly + vbHidden
      On Error GoTo 0
   End If
   RaiseEvent UIStatusSetSub("")
End Sub
Private Sub CopyFileToSoundTable(SourcePath As String, DictId As Long)

   Dim Rs As New Recordset
   
   OpenRecordset Rs, BuildSQL("Select DictId,Sound,SizeInBytes from Sound", "DictId=" & CStr(DictId), "", "")
   If Rs.EOF Then
      Rs.AddNew
      Rs("DictId") = DictId
   End If
   Rs("SizeInBytes") = FileToBlob(Rs("Sound"), SourcePath, Client.SysSettings.SoundChunkSize)
   Rs.Update
   CloseRecordset Rs
End Sub
Private Sub CopyFileFromSoundStorage(DictId As Long, DestPath As String)

   Dim SoundPath As String

   RaiseEvent UIStatusSetSub(Client.Texts.Txt(1260104, "Ljudet hämtas..."))
   If Client.SysSettings.SoundUseImageField Then
      CopyFileFromSoundTable DictId, DestPath
   Else
      SoundPath = GetSoundfilePath(DictId)
      FileCopyWithRetry SoundPath, DestPath
   End If
   RaiseEvent UIStatusSetSub("")
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

Private Sub CopyFileFromSoundTable(DictId As Long, DestPath As String)

   Dim Rs As New Recordset
   
   OpenRecordset Rs, BuildSQL("Select Sound from Sound", "DictId=" & CStr(DictId), "", "")
   If Not Rs.EOF Then
      BlobToFile Rs("Sound"), DestPath, Client.SysSettings.SoundChunkSize
   End If
   CloseRecordset Rs
End Sub
Private Sub CreateSoundSubFolders()
  
   Dim I As Integer
   
   On Error Resume Next
   For I = 0 To 9
      MkDir DictationStorageSoundPath & CStr(I)
      SetAttr DictationStorageSoundPath & CStr(I), vbReadOnly + vbHidden
   Next I
End Sub
Private Function SQLDateString(D As Date) As String

   SQLDateString = SQLDateDelimiter & Format$(D, SQLDateStringFormat) & SQLDateDelimiter
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
Private Function SaveNullZeroLength(S As String, MaxLen As Integer) As Variant

   If Len(S) = 0 Then
      SaveNullZeroLength = Null
   Else
      SaveNullZeroLength = Left$(S, MaxLen)
   End If
End Function
Function FileToBlob(fld As ADODB.Field, Filename As String, Optional ChunkSize As Long = 8192) As Long

   Dim fnum As Integer, BytesLeft As Long, bytes As Long, TotalBytes As Long
   Dim Tmp() As Byte
   Dim T As Double
   Dim TotalTime As Double, KBytePerSecond As Double
    
   T = Timer
   'Can the field handle GetChunk?
   If (fld.Attributes And adFldLong) = 0 Then
       Err.Raise 1001, , "GetChunk can not be used"
   End If
   
   fnum = FreeFile
   Open Filename For Binary As fnum
   TotalBytes = LOF(fnum)
   BytesLeft = TotalBytes
   FileToBlob = TotalBytes
   Do While BytesLeft
       RaiseEvent UIStatusProgress(TotalBytes, BytesLeft)
       bytes = BytesLeft
       If bytes > ChunkSize Then bytes = ChunkSize
       ReDim Tmp(1 To bytes) As Byte
       Get #1, , Tmp
       fld.AppendChunk Tmp
       BytesLeft = BytesLeft - bytes
   Loop
   
   Close #fnum
   TotalTime = Timer - T
   If TotalTime > 0 Then
      KBytePerSecond = TotalBytes / 1024 / TotalTime
   Else
      KBytePerSecond = 0
   End If
   Debug.Print "FileToBlob (" & CStr(TotalBytes) & "/" & Format(KBytePerSecond, "0.0") & ")", Format$(TotalTime, "0.00")

End Function
Sub BlobToFile(fld As ADODB.Field, Filename As String, Optional ChunkSize As Long = 8192)

   Dim fnum As Integer, BytesLeft As Long, bytes As Long, TotalBytes As Long
   Dim TotalTime As Double, KBytePerSecond As Double
   
   Dim Tmp() As Byte
   Dim T As Double
    
   T = Timer
   
   ' Om fältet inte stödjer GetChunk.
   If (fld.Attributes And adFldLong) = 0 Then
       Err.Raise 1001, , "Field doesn't support GetChunk method."
   End If
   KillFileIgnoreError Filename
   
   fnum = FreeFile
   Open Filename For Binary As fnum
   ' Läser filen och skriver till disk.
   TotalBytes = fld.ActualSize
   BytesLeft = TotalBytes
   
   Do While BytesLeft
       RaiseEvent UIStatusProgress(TotalBytes, BytesLeft)
       bytes = BytesLeft
       If bytes > ChunkSize Then bytes = ChunkSize
       Tmp = fld.GetChunk(bytes)
       Put #fnum, , Tmp
       BytesLeft = BytesLeft - bytes
   Loop

   Close #fnum
   
   TotalTime = Timer - T
   If TotalTime > 0 Then
      KBytePerSecond = TotalBytes / 1024 / TotalTime
   Else
      KBytePerSecond = 0
   End If
   Debug.Print "BlobToFile (" & CStr(TotalBytes) & "/" & Format(KBytePerSecond, "0.0") & ")", Format$(TotalTime, "0.00")
End Sub
Public Property Get Connection() As ADODB.Connection

   Set Connection = Db
End Property
