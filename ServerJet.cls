VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServerJet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event UIStatusSet(StatusText As String, Busy As Boolean)
Public Event UIStatusSetSub(SubText As String)
Public Event UIStatusProgress(Total As Long, Left As Long)
Public Event UIStatusClear()

Private Const SettingSection As String = "CareTalkServer"
Private Const DBUserid As String = "GruUser"
Private Const DBPassword As String = "ssdjgfUhssdF"

Private DictationStorageSoundPath As String
Private DictationStoragePath As String
Private SQLDateStringFormat As String
Private SQLDateDelimiter As String
Private mStorageOpened As Boolean

Public Db As ADODB.Connection
Private RsDictList As ADODB.Recordset
Private RsDictTypeList As ADODB.Recordset
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

Public Property Get StorageOpened() As Boolean

   StorageOpened = mStorageOpened
End Property
Public Function IsBatchDoneByClient() As Boolean

   IsBatchDoneByClient = True
End Function

Public Sub DeleteDictations(DeleteOlderThanDays As Integer)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Whr As String
   Dim DeleteOlderThanDate As Date
   
   DeleteOlderThanDate = DateAdd("d", -DeleteOlderThanDays, Now)
   
   Whr = "TranscribedDate<" & SQLDateString(DeleteOlderThanDate) & " and StatusId>=70 and LockedByUserId is null"
   SQL = BuildSQL("Select * from Dictation", Whr, "", "")

   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db

   Do While Not Rs.EOF
      DeleteOneDictation Rs("DictId")
      Rs.MoveNext
   Loop

   CloseRecordset Rs
End Sub
Public Sub DeleteAllDictations()

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from Dictation", "", "", "")

   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db

   Do While Not Rs.EOF
      DeleteOneDictation Rs("DictId")
      Rs.MoveNext
   Loop

   CloseRecordset Rs
End Sub

Private Sub DeleteOneDictation(DictId As Long)
   
   Dim Rs As ADODB.Recordset
   Dim SQL As String

   SQL = BuildSQL("Select * from Dictation", "DictId=" & CStr(DictId), "", "")

   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db

   If Not Rs.EOF Then
      If nvl(Rs("SoundStored"), False) Then
         DeleteSoundFile DictId
      End If
      Rs.Delete
   End If

   CloseRecordset Rs
End Sub
Public Sub DeleteHistory(DeleteOlderThanDays As Integer)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Whr As String
   Dim DeleteOlderThanDate As Date
   
   DeleteOlderThanDate = DateAdd("d", -DeleteOlderThanDays, Now)
   
   Whr = "Created<" & SQLDateString(DeleteOlderThanDate)
   SQL = BuildSQL("Select * from History", Whr, "", "")

   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db

   Do While Not Rs.EOF
      Rs.Delete
      Rs.MoveNext
   Loop

   CloseRecordset Rs
End Sub
Public Sub DeleteSoundFiles(DeleteOlderThanDays As Integer)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Whr As String
   Dim DeleteOlderThanDate As Date
   
   DeleteOlderThanDate = DateAdd("d", -DeleteOlderThanDays, Now)
   
   Whr = "TranscribedDate<" & SQLDateString(DeleteOlderThanDate) & " and SoundStored and StatusId>=70 and StatusId<80 and LockedByUserId is null"
   SQL = BuildSQL("Select * from Dictation", Whr, "", "")

   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db

   Do While Not Rs.EOF
      DeleteSoundFile Rs("DictId")
      Rs("SoundStored") = False
      Rs("SoundDeleted") = Now
      Rs("StatusId") = SoundDeleted
      Rs.MoveNext
   Loop

   CloseRecordset Rs
End Sub
Public Function GetCurrentTimestamp() As Date

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Ts As Double
   
   SQL = BuildSQL("Select LastUsed from TimeStamps", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   Rs.Open SQL, Db
   Ts = CDate(Rs("LastUsed"))
   Rs.Close
   Set Rs = Nothing
   GetCurrentTimestamp = Ts
End Function
Private Sub AddHistory(Dict As clsDict)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Whr As String

   On Error Resume Next
   SQL = BuildSQL("Select * from History", "", "", "")

   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db

   Rs.AddNew
   Rs("DictId") = Dict.DictId
   Rs("Created") = Dict.Created
   Rs("OrgId") = Dict.OrgId
   Rs("DictTypeId") = Dict.DictTypeId
   Rs("PriorityId") = Dict.PriorityId
   Rs("ExpiryDate") = Dict.ExpiryDate
   Rs("AuthorId") = Dict.AuthorId
   Rs("TranscriberId") = Dict.TranscriberId
   Rs("TranscriberOrgId") = Dict.TranscriberOrgId
   Rs("TranscribedDate") = Dict.TranscribedDate
   Rs("SoundLenSec") = Dict.SoundLength
   Rs.Update

   CloseRecordset Rs
   
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
Public Sub SaveGroup(Grp As clsGroup)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If LCase$(Grp.GroupText) = "sysadmin" Then Exit Sub   'Can't change sa group
   
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
   
   SQL = Client.Server.BuildSQL("Select * from Dictation", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   OpenRecordset Rs, SQL
   
   Rs.AddNew
   
   Rs("ExtDictId") = Null
   Rs("ExtSystem") = Null
   
   Rs("LockedByUserId") = Null
   Rs("LockedByStation") = Null
   Rs("LockedTime") = Null
   
   Rs("NoPatient") = Dict.NoPatient
   Rs("PatId") = SaveNullZeroLength(Dict.Pat.PatId, 12)
   Rs("PatName") = SaveNullZeroLength(Dict.Pat.PatName, 50)
   Rs("Created") = Dict.Created
   If Dict.Changed > 0 Then
      Rs("Changed") = Dict.Changed
   End If
   Rs("OrgId") = Dict.OrgId
   Rs("DictTypeId") = Dict.DictTypeId
   Rs("PriorityId") = Dict.PriorityId
   Rs("ExpiryDate") = Dict.ExpiryDate
               
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
   Rs("TimeStamp") = CreateTimeStamp()
   Rs.Update
   If Len(Dict.LocalFilename) > 0 Then
      CopyFileToSoundStorage Dict.LocalFilename, Rs("DictId")
      KillFileIgnoreError Dict.LocalFilename
      Dict.LocalFilename = ""
   End If
   CheckInNewDemo = 0
   Dict.DictId = Rs("DictId")
      
   Rs.Close
   Set Rs = Nothing
End Function
Public Sub CreateDictIdList(OrgId As Long, AuthorId As Long, TranscriberId As Long)

   Dim SQL As String
   
   Set RsDictIdList = New ADODB.Recordset
   If OrgId <> 0 Then
      SQL = BuildSQL("SELECT DictId FROM Dictation", "OrgId=" & CStr(OrgId), "", "DictId")
      RsDictIdList.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
   ElseIf AuthorId <> 0 Then
      SQL = BuildSQL("SELECT DictId FROM Dictation", "AuthorId=" & CStr(AuthorId), "", "DictId")
      RsDictIdList.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
   ElseIf TranscriberId <> 0 Then
      SQL = BuildSQL("SELECT DictId FROM Dictation", "TranscriberId=" & CStr(TranscriberId), "", "DictId")
      RsDictIdList.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
   Else
      SQL = BuildSQL("SELECT DictId FROM Dictation", "", "", "DictId")
      RsDictIdList.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim SelStr As String
   Dim FromStr As String
   Dim WhereStr As String
   Dim GroupByStr As String
   Dim OrderByStr As String
   
   Select Case HistType
      Case htPrio '0
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, PriorityId as RowId, count(DictId) as NumDict, sum(SoundLenSec) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, PriorityId"
      Case htDictType '1
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, DictTypeId as RowId, count(DictId) as NumDict, sum(SoundLenSec) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, DictTypeId"
      Case htOrg '2
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, OrgId as RowId, count(DictId) as NumDict, sum(SoundLenSec) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId"
      Case htTranscriber '3
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, TranscriberId as RowId, count(DictId) as NumDict, sum(SoundLenSec) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, TranscriberId"
      Case htTranscriberOrg  '4
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, TranscriberOrgId as RowId, count(DictId) as NumDict, sum(SoundLenSec) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, TranscriberOrgId"
      Case htDaysFromCreated  '5
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, PriorityId as RowId, sum(Datediff('d',created,TranscribedDate)) as NumDict, count(Dictid) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, PriorityId"
      Case htDaysFromExpiry  '6
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, PriorityId as RowId, sum(Datediff('d',ExpiryDate,TranscribedDate)) as NumDict, count(Dictid) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, PriorityId"
      Case htAuthor '7
         SelStr = "select year(Created) as HistYear, Month(Created) as HistMonth, OrgId, AuthorId as RowId, count(DictId) as NumDict, sum(SoundLenSec) as LenSec "
         FromStr = "FROM History "
         WhereStr = "Year(Created)=" & CStr(HistYear) & " AND OrgId=" & CStr(OrgId) & " "
         GroupByStr = "Year(Created), Month(Created), OrgId, AuthorId"
   End Select

   OpenDisconnectedRecordset RsHist, SelStr & FromStr, WhereStr, GroupByStr, ""
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
   OpenDisconnectedRecordset RsStat, "SELECT * FROM spStatNumber", "OrgId=" & CStr(OrgId), "", ""
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

   NeedLoginInfo = Len(Client.Settings.GetString("DictationStorage", "Cred", "")) = 0
End Function
Public Function UserLogin(ByRef User As clsUser, ByVal LoginName As String, ByVal Password As String, ByVal NewPassword As String) As Integer

   Dim Rs As New ADODB.Recordset
   Dim SQL As String
   Dim RightPassword As String
   
   SQL = BuildSQL("SELECT * FROM Users", "LoginName='" & UCase$(LoginName) & "'", "", "")
   Rs.Open SQL, Db, adOpenDynamic, adLockPessimistic
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
      ElseIf RightPassword = Password Then
         User.UserId = Rs("UserId")
         User.LoginName = Rs("LoginName")
         User.ShortName = Rs("ShortName")
         User.LongName = Rs("LongName")
         User.HomeOrgId = nvl(Rs("HomeOrgid"), 0)
         Rs("LastSuccessfulLogin") = Now
         Rs("LoginFailureCount") = 0
         Rs("UserLockedUntil") = Null
         If Len(NewPassword) > 0 Then
            Rs("Password") = NewPassword
         End If
         UserLogin = 0
      ElseIf Rs("LoginFailureCount") >= Client.SysSettings.LoginMaxRetries Then
         Rs("LoginFailureCount") = 0
         Rs("UserLockedUntil") = DateAdd("n", Client.SysSettings.LoginLockoutMinutes, Now)
         UserLogin = 1
      Else
         Rs("LoginFailureCount") = Rs("LoginFailureCount") + 1
         UserLogin = 2
      End If
      Rs.Update
   Else
      UserLogin = 3
   End If
   Rs.Close
   Set Rs = Nothing
End Function
Public Sub CreateLoggList(StartTime As Date, EndTime As Date, StartLevel As Integer, EndLevel As Integer)

   Dim SQL As String
    
   Set RsLogg = New ADODB.Recordset
   SQL = BuildSQL("Select Logg.*,Users.ShortName as UserShortName from Logg Left join users on Logg.UserId=Users.UserId", _
                  "LoggTime>=" & SQLDateString(StartTime) & " and LoggTime<=" & SQLDateString(EndTime) & " and LoggLevel>=" & CStr(StartLevel) & " and LoggLevel<=" & CStr(EndLevel), _
                  "", _
                  "loggtime desc")
   RsLogg.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim Rs As New ADODB.Recordset
   
   Rs.Open "Logg", Db, adOpenDynamic, adLockOptimistic
   Rs.AddNew
   Rs("LoggTime") = Now
   Rs("LoggId") = Logg.LoggId
   Rs("LoggLevel") = Logg.LoggLevel
   Rs("UserId") = SaveNullIfZero(Logg.UserId)
   Rs("StationId") = SaveNullZeroLength(Logg.StationId, 50)
   Rs("DictId") = SaveNullIfZero(Logg.DictId)
   Rs("LoggData") = SaveNullZeroLength(Logg.LoggData, 50)
   Rs.Update
   Rs.Close
   Set Rs = Nothing
End Sub
Public Sub CreateDictAuditList(DictId As Long)

   Dim SQL As String
   
   SQL = BuildSQL("Select DictAudit.*,Users.ShortName as UserShortName from DictAudit Left join users on DictAudit.UserId=Users.UserId", _
                "DictId=" & CStr(DictId), _
                "", _
                "AuditTime desc")
   
   Set RsDictAudit = New ADODB.Recordset
   RsDictAudit.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim SQL As String
   
   SQL = BuildSQL("Select* from ExtSystem", _
                "", _
                "", _
                "")
   
   Set RsExtSystem = New ADODB.Recordset
   RsExtSystem.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Set E = New clsExtSystem
   With E
      .ExtSystem = Rs("ExtSystem")
   End With
End Sub
Public Sub CreateExtSystemDictTypeList()

   Dim SQL As String
   
   SQL = BuildSQL("Select* from ExtSystemDictType", _
                "", _
                "", _
                "")
   
   Set RsExtSystemDictType = New ADODB.Recordset
   RsExtSystemDictType.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim SQL As String
   
   SQL = BuildSQL("Select* from ExtSystemOrg", _
                "", _
                "", _
                "")
   
   Set RsExtSystemOrg = New ADODB.Recordset
   RsExtSystemOrg.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim SQL As String
   
   SQL = BuildSQL("Select* from ExtSystemPriority", _
                "", _
                "", _
                "")
   
   Set RsExtSystemPriority = New ADODB.Recordset
   RsExtSystemPriority.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim Rs As New ADODB.Recordset
   
   Rs.Open "DictAudit", Db, adOpenDynamic, adLockOptimistic
   Rs.AddNew
   Rs("DictId") = Audit.DictId
   Rs("AuditTime") = Now
   Rs("AuditType") = Audit.AuditType
   Rs("DictStatus") = Audit.DictStatus
   Rs("UserId") = SaveNullIfZero(Audit.UserId)
   Rs("StationId") = SaveNullZeroLength(Audit.StationId, 50)
   Rs.Update
   Rs.Close
   Set Rs = Nothing
End Sub
Public Sub CreateOrgList()

   OpenDisconnectedRecordset RsOrgList, "SELECT * FROM Org", "", "", ""
   Set RsOrgListClone = RsOrgList.Clone
   OpenDisconnectedRecordset RsUserRole, "SELECT * FROM spUserRole", "Userid=" & Client.User.UserId, "", ""
   
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
   Dim R As clsRoles
   
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

   Dim R As clsRoles
   Dim Sum As clsRoles
   
   If Not RsUserRole.EOF Or Not RsUserRole.BOF Then
      RsUserRole.MoveFirst
      Do While Not RsUserRole.EOF
         If nvl(RsUserRole("Orgid"), 0) = OrgId Then
            Set R = New clsRoles
            
            R.OrgId = RsUserRole("OrgId")
            R.Roles = RsUserRole("Roles")
            R.Delayed = RsUserRole("Delayed")
            R.DelayedHours = RsUserRole("DelayedHours")
            
            AddRole Res, R
         End If
         RsUserRole.MoveNext
      Loop
   End If
End Sub
Private Sub GetRolesFromOrgParent(Res As clsRoles, ByVal Id As Long)

   Dim R As clsRoles
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

Public Sub CreateDictList(ByVal OrgId As Long, ByVal AuthorId As Long, ByVal TranscriberId As Long, StatusIdStart As Long, StatusIdEnd As Long, TimeStamp As Double)

   Dim WhereClause As String
   Dim SelectClause As String
      
   If OrgId > 0 Then
      AddString WhereClause, "AND", "OrgId=" & CStr(OrgId)
   End If
   If AuthorId > 0 Then
      AddString WhereClause, "AND", "AuthorId=" & CStr(AuthorId)
   End If
   If TranscriberId > 0 Then
      AddString WhereClause, "AND", "TranscriberId=" & CStr(TranscriberId)
   End If
   If StatusIdStart > 0 Then
      AddString WhereClause, "AND", "StatusId>=" & CStr(StatusIdStart)
   End If
   If StatusIdEnd > 0 Then
      AddString WhereClause, "AND", "StatusId<=" & CStr(StatusIdEnd)
   End If

   AddString WhereClause, "AND", "TimeStamp>" & Format$(TimeStamp)
   
   OpenDisconnectedRecordset RsDictList, "SELECT * FROM spDictation", WhereClause, "", "DictId ASC"
End Sub
Public Function DictListGetNext(ByRef Dict As clsDict) As Boolean

   If RsDictList.EOF Then
      DictListGetNext = False
      RsDictList.Close
      Set RsDictList = Nothing
   Else
      ConvertDictRsToClass RsDictList, Dict
      RsDictList.MoveNext
      DictListGetNext = True
   End If
End Function
Private Sub ConvertDictRsToClass(ByRef Rs As ADODB.Recordset, ByRef Dict As clsDict)

   Set Dict = New clsDict
   With Dict
      .DictId = Rs("DictId")
      .ExtDictId = nvl(Rs("ExtDictId"), 0)
      .ExtSystem = nvl(Rs("ExtSystem"), "")
      .NoPatient = nvl(Rs("NoPatient"), False)
      .Pat.PatId = nvl(Rs("PatId"), "")
      .Pat.PatId2 = nvl(Rs("PatId2"), "")
      .Pat.PatName = nvl(Rs("PatName"), "")
      .Created = Rs("Created")
      .Changed = nvl(Rs("Changed"), 0)
      .SoundLength = nvl(Rs("SoundLength"), 0)
      .OrgId = Rs("OrgId")
      .OrgText = Rs("OrgText")
      .DictTypeId = Rs("DictTypeId")
      .DictTypeText = Rs("DictTypeText")
      .AuthorId = Rs("AuthorId")
      .AuthorShortName = Rs("AuthorShortName")
      .AuthorLongName = Rs("AuthorLongName")
      .TranscriberId = nvl(Rs("TranscriberId"), 0)
      .TranscriberOrgId = nvl(Rs("TranscriberOrgId"), 0)
      .TranscribedDate = nvl(Rs("TranscribedDate"), 0)
      .TranscriberShortName = nvl(Rs("TranscriberShortName"), "")
      .TranscriberLongName = nvl(Rs("TranscriberLongName"), "")
      .StatusId = Rs("StatusId")
      .StatusText = Rs("StatusText")
      .PriorityId = Rs("PriorityId")
      .PriorityText = Rs("PriorityText")
      .ExpiryDate = Rs("ExpiryDate")
      .Txt = nvl(Rs("Txt"), "")
      .LockedByStation = nvl(Rs("LockedByStation"), "")
      .LockedByUserShortName = nvl(Rs("LockedByUserShortName"), "")
      .LockedByUserLongName = nvl(Rs("LockedByUserLongName"), "")
      .LockedTime = nvl(Rs("LockedTime"), 0)
      .TimeStamp = Rs("Timestamp")
      .SoundDeleted = nvl(Rs("SoundDeleted"), 0)
      
      .SoundDirty = False
      .InfoDirty = False
   End With
End Sub
Public Sub CreateDictTypeList()

   OpenDisconnectedRecordset RsDictTypeList, "SELECT * FROM DictType", "", "", "DictTypeText ASC"
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

   With DictType
      .DictTypeId = Rs("DictTypeId")
      .DictTypeText = Rs("DictTypeText")
   End With
End Sub
Public Sub CreateGroupList()

   OpenDisconnectedRecordset RsGroupList, "SELECT * FROM Groups", "", "", "GroupText ASC"
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

   With Grp
      .GroupId = Rs("GroupId")
      .GroupText = Rs("GroupText")
      .GroupDesc = nvl(Rs("GroupDesc"), "")
      .AdmOrgId = nvl(Rs("AdmOrgId"), 0)
   End With
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
      RsUsers.Close
      Set RsUsers = Nothing
   Else
      ConvertUsersRsToClass RsUsers, Usr
      RsUsers.MoveNext
      UserListGetNext = True
   End If
End Function
Private Sub ConvertUsersRsToClass(ByRef Rs As ADODB.Recordset, ByRef Usr As clsUser)

   Set Usr = New clsUser

   With Usr
      .UserId = Rs("UserId")
      .LoginName = Rs("LoginName")
      .Password = Rs("Password")
      .ShortName = Rs("ShortName")
      .LongName = Rs("LongName")
      .HomeOrgId = nvl(Rs("HomeOrgId"), 0)
   End With
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

   With Priority
      .PriorityId = Rs("PriorityId")
      .PriortyText = Rs("PriorityText")
      .Days = Rs("Days")
   End With
End Sub
Private Sub Class_Terminate()

   DictationStorageClose
   
End Sub
Public Function DictationStorageOpen(UserId As String, Password As String) As Boolean

   Dim Rs As New ADODB.Recordset
   Dim CS As String
   Dim Provider As String
   Dim SystemDb As String
   Dim DbFile As String
   Dim CurrUserid As String
   Dim CurrPassword As String

   On Error GoTo DictationStorageOpen_Err
   
   SQLDateStringFormat = Client.Settings.GetString("DictationStorage", "SQLDateStringFormat", "yyyy-mm-dd")
   SQLDateDelimiter = Client.Settings.GetString("DictationStorage", "SQLDateDelimiter", "#")
   
   DictationStoragePath = Client.Settings.GetFolder("DictationStorage", "Folder", "C:\CareTalk\Storage\")
   Provider = Client.Settings.GetString("DictationStorage", "Provider", "Microsoft.Jet.OLEDB.4.0")
   SystemDb = Client.Settings.GetString("DictationStorage", "SystemDB", DictationStoragePath & "Dictation.cdw")
   DbFile = Client.Settings.GetString("DictationStorage", "DbFile", DictationStoragePath & "Dictation.cdb")
   CurrUserid = Client.Settings.GetString("DictationStorage", "DbUid", DBUserid)
   CurrPassword = Client.Settings.GetString("DictationStorage", "DbPw", DBPassword)

   Set Db = New ADODB.Connection
   CS = "Provider=" & Provider
   CS = CS & ";jet oledb:system database='" & SystemDb & "'"
   CS = CS & ";Data source='" & DbFile & "'"
   CS = CS & ";User Id=" & CurrUserid & ";Password=" & CurrPassword
   
   Db.ConnectionString = CS
   Db.Open
   
   DictationStorageSoundPath = Client.Settings.GetFolder("DictationStorage", "SoundFolder", DictationStoragePath)
   CreateSoundSubFolders
   mStorageOpened = True
   DictationStorageOpen = mStorageOpened
   Exit Function
   
DictationStorageOpen_Err:
   mStorageOpened = False
   DictationStorageOpen = mStorageOpened
   Exit Function
End Function

Private Sub DictationStorageClose()

   On Error Resume Next
   mStorageOpened = False
   Db.Close
   Set Db = Nothing
   
   DictationStorageSoundPath = ""
End Sub

Public Function CheckOutDict(ByRef Dict As clsDict, ByVal DictId As Long, ByVal WithSound As Boolean, ByVal LocalTempFileName As String) As Integer

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   RaiseEvent UIStatusSet(Client.Texts.Txt(1260101, "Diktatet hämtas"), True)
   If LockDictation(DictId) Then
      
      SQL = BuildSQL("Select * from spDictation", "DictId=" & CStr(DictId), "", "")
      
      Set Rs = New ADODB.Recordset
      Rs.CursorLocation = adUseClient
      Rs.CursorType = adOpenDynamic
      Rs.LockType = adLockPessimistic
      Rs.Open SQL, Db
      
      If Rs.EOF Then
         CheckOutDict = 1
      Else
         If nvl(Rs("LockedByUserId"), 0) <> Client.User.UserId Then
            CheckOutDict = 2
         Else
            ConvertDictRsToClass Rs, Dict
            If WithSound And Len(LocalTempFileName) > 0 And Dict.StatusId < SoundDeleted Then
               Dict.LocalFilename = LocalTempFileName
               CopyFileFromSoundStorage DictId, LocalTempFileName
            End If
            
            Dict.ReadOnly = Dict.AuthorId <> Client.User.UserId Or Dict.StatusId > 30
            
            CheckOutDict = 0
         End If
         Rs.Close
         Set Rs = Nothing
      End If
   Else
      CheckOutDict = 2
   End If
   RaiseEvent UIStatusClear
End Function
Public Function GetDictIdFromExtDictId(ExtSystem As String, ExtDictId As String) As Long

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   If Len(ExtSystem) = 0 Or Len(ExtDictId) = 0 Then Exit Function
   
   SQL = BuildSQL("Select * from spDictation", "ExtDictId='" & ExtDictId & "' and ExtSystem='" & ExtSystem & "'", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   
   If Rs.EOF Then
      GetDictIdFromExtDictId = 0
   Else
      GetDictIdFromExtDictId = Rs!DictId
   End If
   Rs.Close
   Set Rs = Nothing
End Function
Private Function LockDictation(DictId As Long) As Boolean

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   On Error GoTo LockDictation_Err
   SQL = BuildSQL("Select * from Dictation", "DictId=" & CStr(DictId), "", "")
      
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   If nvl(Rs("LockedByUserId"), 0) = 0 Or (nvl(Rs("LockedByUserId"), 0) = Client.User.UserId And nvl(Rs("LockedByStation"), "") = Client.Station.Id) Then
      Rs("LockedByUserId") = Client.User.UserId
      Rs("LockedByStation") = Client.Station.Id
      Rs("LockedTime") = Now
      Rs("TimeStamp") = CreateTimeStamp()
      Rs.Update
      LockDictation = True
   End If
   CloseRecordset Rs
   Exit Function
   
LockDictation_Err:
   CloseRecordset Rs
   LockDictation = False
   Exit Function
End Function
Private Sub UnLockDictationWithNewStatus(DictId As Long, StatusId As Long)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   On Error GoTo UnLockDictationWithNewStatus_Err
   SQL = BuildSQL("Select * from Dictation", "DictId=" & CStr(DictId), "", "")
      
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   If Not Rs.EOF Then
      Rs("StatusId") = StatusId
      Rs("LockedByUserId") = Null
      Rs("LockedByStation") = Null
      Rs("LockedTime") = Null
      Rs("TimeStamp") = CreateTimeStamp()
      Rs.Update
   End If
   CloseRecordset Rs
   Exit Sub
   
UnLockDictationWithNewStatus_Err:
   CloseRecordset Rs
   Exit Sub
End Sub
Public Function CheckInDict(ByRef Dict As clsDict, DiscardChanges As Boolean)

   Dim SQL As String
   Dim Rs As ADODB.Recordset
   Dim Dirty As Boolean
   
   RaiseEvent UIStatusSet(Client.Texts.Txt(1260102, "Diktatet sparas"), True)

   SQL = BuildSQL("Select * from Dictation", "DictId=" & CStr(Dict.DictId), "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   
   If Rs.EOF Then
      CheckInDict = 1
   Else
      If Rs("LockedByUserId") <> Client.User.UserId Then
         CheckInDict = 2
      Else
         Rs("LockedByUserId") = Null
         Rs("LockedByStation") = Null
         Rs("LockedTime") = Null
         
         If Not DiscardChanges Then
            If Rs("AuthorId") = Client.User.UserId Then
               If Rs("NoPatient") <> Dict.NoPatient Then
                  Rs("NoPatient") = SaveBool(Dict.NoPatient)
                  Dirty = True
               End If
               If Rs("PatId") <> Dict.Pat.PatId Then
                  Rs("PatId") = SaveNullZeroLength(Dict.Pat.PatId, 12)
                  Dirty = True
               End If
               If Rs("PatId2") <> Dict.Pat.PatId2 Then
                  Rs("PatId2") = SaveNullZeroLength(Dict.Pat.PatId2, 12)
                  Dirty = True
               End If
               If Rs("PatName") <> Dict.Pat.PatName Then
                  Rs("PatName") = SaveNullZeroLength(Dict.Pat.PatName, 50)
                  Dirty = True
               End If
               If Rs("OrgId") <> Dict.OrgId Then
                  Rs("OrgId") = Dict.OrgId
                  Dirty = True
               End If
               If Rs("DictTypeId") <> Dict.DictTypeId Then
                  Rs("DictTypeId") = Dict.DictTypeId
                  Dirty = True
               End If
               If Rs("PriorityId") <> Dict.PriorityId Then
                  Rs("PriorityId") = Dict.PriorityId
                  Dirty = True
               End If
               If Rs("ExpiryDate") <> Dict.ExpiryDate Then
                  Rs("ExpiryDate") = Dict.ExpiryDate
                  Dirty = True
               End If
               If nvl(Rs("Txt"), "") <> Dict.Txt Then
                  Rs("Txt") = SaveNullZeroLength(Dict.Txt, 50)
                  Dirty = True
               End If
               
               If Len(Dict.LocalFilename) > 0 Then
                  If Dict.SoundDirty Then
                     CopyFileToSoundStorage Dict.LocalFilename, Dict.DictId
                     Rs("SoundChanged") = Now
                  End If
               End If
               
               If Dirty Or Dict.SoundDirty Then
                  Dict.Changed = Now
               End If
            End If
            If Rs("StatusId") < BeingTrancribed And Dict.StatusId >= BeingTrancribed And Dict.StatusId < SoundDeleted Then
               Rs("TranscriberId") = Client.User.UserId
               If Client.User.HomeOrgId > 0 Then
                  Rs("TranscriberOrgId") = Client.User.HomeOrgId
               End If
            End If
            If Rs("StatusId") < WaitForSign And Dict.StatusId >= WaitForSign And Dict.StatusId < SoundDeleted Then
               Dict.TranscriberId = Client.User.UserId
               Rs("TranscriberId") = Dict.TranscriberId
               If Client.User.HomeOrgId > 0 Then
                  Dict.TranscriberOrgId = Client.User.HomeOrgId
                  Rs("TranscriberOrgId") = Dict.TranscriberOrgId
               End If
               Dict.TranscribedDate = Now
               Rs("TranscribedDate") = Dict.TranscribedDate
               AddHistory Dict
            End If
            Rs("StatusId") = Dict.StatusId
            
            If Len(Dict.LocalFilename) > 0 Then
               Rs("SoundStored") = True
               Rs("SoundLength") = Dict.SoundLength
            End If
         End If
         Rs("TimeStamp") = CreateTimeStamp()
         Rs.Update
         CheckInDict = 0
      End If
   End If
   Rs.Close
   Set Rs = Nothing
   RaiseEvent UIStatusClear
End Function
Public Function CheckInNewDict(ByRef Dict As clsDict) As Integer

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Dirty As Boolean
   
   RaiseEvent UIStatusSet(Client.Texts.Txt(1260102, "Diktatet sparas"), True)
   SQL = BuildSQL("Select * from Dictation", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   
   Rs.AddNew
   
   Rs("ExtDictId") = SaveNullZeroLength(Dict.ExtDictId, 20)
   Rs("ExtSystem") = SaveNullZeroLength(Dict.ExtSystem, 10)
   
   Rs("LockedByUserId") = Null
   Rs("LockedByStation") = Null
   Rs("LockedTime") = Null
   
   Rs("NoPatient") = SaveBool(Dict.NoPatient)
   Rs("PatId") = SaveNullZeroLength(Dict.Pat.PatId, 12)
   Rs("PatId2") = SaveNullZeroLength(Dict.Pat.PatId2, 12)
   Rs("PatName") = SaveNullZeroLength(Dict.Pat.PatName, 50)
   Rs("Created") = Now
   Rs("OrgId") = Dict.OrgId
   Rs("DictTypeId") = Dict.DictTypeId
   Rs("PriorityId") = Dict.PriorityId
   Rs("ExpiryDate") = Dict.ExpiryDate
   Rs("Txt") = SaveNullZeroLength(Dict.Txt, 50)
   
   If Len(Dict.LocalFilename) > 0 Then
      Rs("SoundStored") = True
      Rs("SoundCreated") = Now
      Rs("SoundLength") = Dict.SoundLength
   End If
               
   Rs("StatusId") = Dict.StatusId
   Rs("AuthorId") = Client.User.UserId
   Rs("TranscriberId") = Null
   Rs("TranscriberOrgId") = Null
   Rs("TranscribedDate") = Null
   Rs("SoundDeleted") = Null
   Rs("TimeStamp") = CreateTimeStamp()
   Rs.Update
   If Len(Dict.LocalFilename) > 0 Then
      CopyFileToSoundStorage Dict.LocalFilename, Rs("DictId")
      KillFileIgnoreError Dict.LocalFilename
      Dict.LocalFilename = ""
   End If
   CheckInNewDict = 0
   Dict.DictId = Rs("DictId")
   Rs.Close
   Set Rs = Nothing
   RaiseEvent UIStatusClear
End Function
Public Sub UnlockDict(DictId As Long)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   On Error GoTo UnLockDict_Err
   SQL = BuildSQL("Select * from Dictation", "DictId=" & CStr(DictId), "", "")
      
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   If Not Rs.EOF Then
      Rs("LockedByUserId") = Null
      Rs("LockedByStation") = Null
      Rs("LockedTime") = Null
      Rs("TimeStamp") = CreateTimeStamp()
      Rs.Update
   End If
   CloseRecordset Rs
   Exit Sub
   
UnLockDict_Err:
   CloseRecordset Rs
   Exit Sub
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
      Letter = Mid$(R2, I, 1)
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
Public Sub OpenDisconnectedRecordset(ByRef Rs As ADODB.Recordset, ByVal SelectFrom As String, ByVal WhereClause As String, ByVal GroupByClause As String, ByVal OrderClause As String)

   Dim SQL As String

   CloseRecordset Rs
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenStatic
   Rs.LockType = adLockReadOnly
   SQL = BuildSQL(SelectFrom, WhereClause, GroupByClause, OrderClause)
   Rs.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
   Rs.ActiveConnection = Nothing
End Sub
Private Sub OpenRecordset(ByRef Rs As ADODB.Recordset, ByRef SQL As String)

   Rs.Open SQL, Db, adOpenDynamic, adLockOptimistic
End Sub
Public Sub CloseRecordset(Rs As ADODB.Recordset)

   On Error Resume Next
   Rs.Close
   Set Rs = Nothing
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
Private Function CreateTimeStamp() As Double

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   Dim Ts As Double
   
   SQL = BuildSQL("Select LastUsed from TimeStamps", "", "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   Ts = Rs("LastUsed") + 1
   Rs("LastUsed") = Ts
   Rs.Update
   Rs.Close
   Set Rs = Nothing
   CreateTimeStamp = Ts
End Function
Public Sub WriteUserData(Ty As String, Ke As String, Va As String)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from UserData", "Ty='" & Ty & "' and Ke='" & Ke & "' and UserId=" & Client.User.UserId, "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   If Rs.EOF Then
      Rs.AddNew
      Rs("Userid") = Client.User.UserId
      Rs("Ty") = Ty
      Rs("Ke") = Ke
   End If
   Rs("Va") = Va
   Rs("Version") = ApplicationVersion
   Rs.Update
   Rs.Close
   Set Rs = Nothing
End Sub
Public Function ReadUserData(Ty As String, Ke As String, Def As String, ByRef Version) As String

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from UserData", "Ty='" & Ty & "' and Ke='" & Ke & "' and UserId=" & Client.User.UserId, "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   Rs.Open SQL, Db
   If Rs.EOF Then
      ReadUserData = Def
      Version = ""
   Else
      ReadUserData = Rs("Va")
      Version = Rs("Version")
   End If
   Rs.Close
   Set Rs = Nothing
End Function
Public Sub CreateSysSettingsList(ByVal Ty As String)

   Dim SQL As String
   
   SQL = BuildSQL("select * from SysSettings", "ty='" & Ty & "'", "", "Ty, Se, Ke")
   
   Set RsSysSettingsList = New ADODB.Recordset
   RsSysSettingsList.Open SQL, Db, adOpenForwardOnly, adLockReadOnly
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

   Dim Rs As New ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from SysSettings", "Ty='" & Ty & "' and Se='" & Se & "' and Ke='" & Ke & "'", "", "")
   Rs.Open SQL, Db, adOpenDynamic, adLockPessimistic
   If Rs.EOF Then
      Rs.AddNew
      Rs("Ty") = Ty
      Rs("Se") = Se
      Rs("Ke") = Ke
   End If
   Rs("Va") = S
   Rs.Update
   Rs.Close
End Function
Public Sub WriteStationData(Ty As String, Ke As String, Va As String)

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from StationData", "Ty='" & Ty & "' and Ke='" & Ke & "' and UserId=" & Client.User.UserId, "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenDynamic
   Rs.LockType = adLockPessimistic
   Rs.Open SQL, Db
   If Rs.EOF Then
      Rs.AddNew
      Rs("Stationid") = Client.Station.Id
      Rs("Ty") = Ty
      Rs("Ke") = Ke
   End If
   Rs("Va") = Va
   Rs("Version") = ApplicationVersion
   Rs.Update
   Rs.Close
   Set Rs = Nothing
End Sub
Public Function ReadStationData(Ty As String, Ke As String, Def As String, ByRef Version) As String

   Dim Rs As ADODB.Recordset
   Dim SQL As String
   
   SQL = BuildSQL("Select * from StationData", "Ty='" & Ty & "' and Ke='" & Ke & "' and UserId=" & Client.User.UserId, "", "")
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockReadOnly
   Rs.Open SQL, Db
   If Rs.EOF Then
      ReadStationData = Def
      Version = ""
   Else
      ReadStationData = Rs("Va")
      Version = Rs("Version")
   End If
   Rs.Close
   Set Rs = Nothing
End Function

Private Sub CopyFileToSoundStorage(SourcePath As String, DictId As Long)

   Dim SoundPath As String

   RaiseEvent UIStatusSetSub(Client.Texts.Txt(1260103, "Ljudet sparas..."))

   SoundPath = GetSoundfilePath(DictId)
   On Error Resume Next
   SetAttr SoundPath, vbNormal
   On Error GoTo 0
   FileCopy SourcePath, SoundPath
   SetAttr SoundPath, vbReadOnly + vbHidden
   RaiseEvent UIStatusSetSub("")
End Sub
Private Sub CopyFileFromSoundStorage(DictId As Long, DestPath As String)

   Dim SoundPath As String

   RaiseEvent UIStatusSetSub(Client.Texts.Txt(1260104, "Ljudet hämtas..."))
   SoundPath = GetSoundfilePath(DictId)
   FileCopy SoundPath, DestPath
   RaiseEvent UIStatusSetSub("")
End Sub
Private Sub DeleteSoundFile(DictId As Long)

   Dim SoundPath As String

   On Error Resume Next
   SoundPath = GetSoundfilePath(DictId)
   SetAttr SoundPath, vbNormal
   Kill SoundPath
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
Private Function SaveBool(B As Boolean) As Boolean

   If B Then
      SaveBool = True
   Else
      SaveBool = False
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

