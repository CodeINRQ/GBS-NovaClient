VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TempDictInfoTypeEnum
   tdiUpdateInfo = -1
   tdiEmpty = 0
   tdiNew = 1
End Enum

Private Const MaxNumberDictInSearchResult = 500
Private DictCache(MaxNumberOfDictation) As clsDict
Private DictCacheNumber As Integer
Private DictCacheIndex As Integer
Private DictCacheTimeStampTot As Double
Private DictCacheTimeStampOrg(30050) As Double

Private DictIdList(MaxNumberOfDictation) As Long
Private DictIdListIdx As Integer

Private DictListTimeStamp As Double
Private FilterCol As Collection
Private mSearchFilter As clsFilter
Private mCurrPatientFilter As clsFilter
Private LastCurrPatId As String
Private LastCurrPatId2 As String

Public Function GetDictFromCache(DictId As Long, Dict As clsDict) As Boolean

   Dim L As Long
   
   For L = 0 To DictCacheNumber - 1
      If Not IsNull(DictCache(L)) Then
         If DictCache(L).DictId = DictId Then
            Set Dict = DictCache(L)
            GetDictFromCache = True
         End If
      End If
   Next L
End Function
Public Sub UnlockDict(DictId As Long)

   Dim Dict As clsDict
   
   Client.Server.UnlockDict DictId
   GetDictFromCache DictId, Dict
   Client.LoggMgr.Insert 1320117, LoggLevel_DictAdmin, Dict.DictId, Dict.LoggData
   Client.DictAuditMgr.Insert DictId, AuditType_Unlocked, Dict.StatusId
End Sub
Public Property Set SearchFilter(Flt As clsFilter)

   DictCacheTimeStampOrg(30050) = -1
   Set mSearchFilter = Flt
End Property
Public Property Set CurrPatientFilter(Flt As clsFilter)

   DictCacheTimeStampOrg(30005) = -1
   Set mCurrPatientFilter = Flt
End Property

Public Function CreateList(ByVal OrgId As Long, TimeStamp As Double, ByRef TooMany As Boolean) As Double

   DictListTimeStamp = TimeStamp
   
   Set FilterCol = Nothing
   Set FilterCol = New Collection
   Select Case OrgId
      Case Is < 30000
         BuildFilterForOrgId OrgId
      Case 30000 'My dictations
         AddFilter 0, Client.User.UserId, 0, BeingRecorded, Transcribed
         AddFilter 0, 0, Client.User.UserId, BeingRecorded, Transcribed
         'AddFilter 0, Client.User.UserId, 0, BeingRecorded, BeingRecorded
         'AddFilter 0, Client.User.UserId, 0, Recorded, Recorded
         'AddFilter 0, 0, Client.User.UserId, BeingTrancribed, BeingTrancribed
         'AddFilter 0, Client.User.UserId, 0, BeingTrancribed, BeingTrancribed
         'If Client.SysSettings.UseAuthorsSign Then
         '   AddFilter 0, 0, Client.User.UserId, WaitForSign, WaitForSign
         '   AddFilter 0, Client.User.UserId, 0, WaitForSign, WaitForSign
         'End If
         'AddFilter 0, 0, Client.User.UserId, Transcribed, Transcribed
         'AddFilter 0, Client.User.UserId, 0, Transcribed, Transcribed
      Case 30010 'Being recorded
         AddFilter 0, Client.User.UserId, 0, BeingRecorded, BeingRecorded
      Case 30005 'Current patient
         If Not mCurrPatientFilter Is Nothing Then
            If Len(mCurrPatientFilter.Pat.PatId) > 0 Or Len(mCurrPatientFilter.Pat.PatId2) > 0 Then
               FilterCol.Add mCurrPatientFilter
            End If
         End If
      Case 30020 'Recorded
         AddFilter 0, Client.User.UserId, 0, Recorded, Recorded
      Case 30025 'Being transcribed
         AddFilter 0, 0, Client.User.UserId, BeingTrancribed, BeingTrancribed
         AddFilter 0, Client.User.UserId, 0, BeingTrancribed, BeingTrancribed
      Case 30030 'waiting for sign
         If Client.SysSettings.UseAuthorsSign Then
            AddFilter 0, 0, Client.User.UserId, WaitForSign, WaitForSign
            AddFilter 0, Client.User.UserId, 0, WaitForSign, WaitForSign
         End If
      Case 30040 'Transcribed
         AddFilter 0, 0, Client.User.UserId, Transcribed, Transcribed
         AddFilter 0, Client.User.UserId, 0, Transcribed, Transcribed
      Case 30050 'Sök
         If Not mSearchFilter Is Nothing Then
            FilterCol.Add mSearchFilter
         End If
   End Select
   
   DictCacheIndex = 0
   CreateList = FillCache(OrgId, TooMany)
   Debug.Print "DictCacheNumber " & CStr(DictCacheNumber)
End Function

Public Function ListNextItem(ByRef Dict As clsDict) As Boolean

   Dim Hit As Boolean

   Do While DictCacheIndex < DictCacheNumber
      If FilterDictation(DictCache(DictCacheIndex)) Then
         Hit = True
         Set Dict = DictCache(DictCacheIndex)
      End If
      DictCacheIndex = DictCacheIndex + 1
      If Hit Then
         ListNextItem = True
         Exit Function
      End If
   Loop
   ListNextItem = False
End Function

Private Function FillCache(ByVal OrgId As Long, ByRef TooMany As Boolean) As Double

   Dim Dict As clsDict
   Dim MaxTimeStamp As Double
   Static CleanCounter As Integer
   Dim FirstUpdate As Boolean

   If DictCacheTimeStampTot = 0 Then
      DictCacheTimeStampTot = Client.Server.GetCurrentTimestamp()
   End If

   If DictCacheTimeStampOrg(OrgId) < 0 Then
      DictCacheTimeStampOrg(OrgId) = FillCacheFromFilter(OrgId, TooMany)
      FillCache = DictCacheTimeStampOrg(OrgId)
   Else
      Client.Server.CreateDictList 0, 0, 0, 0, 255, DictCacheTimeStampTot
      MaxTimeStamp = DictCacheTimeStampTot
      FirstUpdate = True
      Do While Client.Server.DictListGetNext(Dict)
         If MaxTimeStamp < Dict.TimeStamp Then
            MaxTimeStamp = Dict.TimeStamp
         End If
         UpdateOneDictInCache Dict, FirstUpdate
         FirstUpdate = False
      Loop
      DictCacheTimeStampTot = MaxTimeStamp
      FillCache = MaxTimeStamp
      
      CleanCounter = CleanCounter + 1
      If CleanCounter > 10 Then
         CleanCounter = 0
         CleanCacheFromFilter OrgId
      End If
   End If
End Function
Private Sub CleanCacheFromFilter(OrgId As Long)

   Dim Flt As clsFilter

   For Each Flt In FilterCol
      If Flt.OrgId <> 0 Then
         CleanCacheFromOrgId Flt.OrgId
      ElseIf Flt.AuthorId <> 0 Then
         CleanCacheFromAuthorId Flt.AuthorId
      ElseIf Flt.TranscriberId <> 0 Then
         CleanCacheFromTranscriberId Flt.TranscriberId
      ElseIf OrgId = 30005 Then
         CleanCacheFromPatId mCurrPatientFilter.Pat.PatId, mCurrPatientFilter.Pat.PatId2
      'Else
      '   CleanCacheAll
      End If
   Next Flt
End Sub
Private Sub CleanCacheFromOrgId(OrgId As Long)

   Dim I As Integer

   DictIdListIdx = 0
   Client.Server.CreateDictIdList OrgId, 0, 0
   Do While Client.Server.DictIdGetNext(DictIdList(DictIdListIdx))
      If DictIdListIdx < UBound(DictIdList) Then
         DictIdListIdx = DictIdListIdx + 1
      End If
   Loop
   
   For I = 0 To DictCacheNumber - 1
      If DictCache(I).OrgId = OrgId Then
         If Not IsDictIdInDictIdList(DictCache(I).DictId) Then
            DeleteDictFromCache I
         End If
      End If
   Next I
End Sub
Private Sub CleanCacheFromAuthorId(AuthorId As Long)

   Dim I As Integer

   DictIdListIdx = 0
   Client.Server.CreateDictIdList 0, AuthorId, 0
   Do While Client.Server.DictIdGetNext(DictIdList(DictIdListIdx))
      If DictIdListIdx < UBound(DictIdList) Then
         DictIdListIdx = DictIdListIdx + 1
      End If
   Loop
   
   For I = 0 To DictCacheNumber - 1
      If DictCache(I).AuthorId = AuthorId Then
         If Not IsDictIdInDictIdList(DictCache(I).DictId) Then
            DeleteDictFromCache I
         End If
      End If
   Next I
End Sub
Private Sub CleanCacheFromTranscriberId(TranscriberId As Long)

   Dim I As Integer

   DictIdListIdx = 0
   Client.Server.CreateDictIdList 0, 0, TranscriberId
   Do While Client.Server.DictIdGetNext(DictIdList(DictIdListIdx))
      If DictIdListIdx < UBound(DictIdList) Then
         DictIdListIdx = DictIdListIdx + 1
      End If
   Loop
   
   For I = 0 To DictCacheNumber - 1
      If DictCache(I).TranscriberId = TranscriberId Then
         If Not IsDictIdInDictIdList(DictCache(I).DictId) Then
            DeleteDictFromCache I
         End If
      End If
   Next I
End Sub
Private Sub CleanCacheFromPatId(PatId As String, PatId2 As String)

   Dim I As Integer

   DictIdListIdx = 0
   Client.Server.CreateDictIdList 0, 0, 0, PatId, PatId2
   Do While Client.Server.DictIdGetNext(DictIdList(DictIdListIdx))
      If DictIdListIdx < UBound(DictIdList) Then
         DictIdListIdx = DictIdListIdx + 1
      End If
   Loop
   
   For I = 0 To DictCacheNumber - 1
      If Len(DictCache(I).Pat.PatId) > 0 Then
         If (DictCache(I).Pat.PatId = PatId Or DictCache(I).Pat.PatId = PatId2) Then
            If Not IsDictIdInDictIdList(DictCache(I).DictId) Then
               DeleteDictFromCache I
            End If
         End If
      End If
   Next I
End Sub
Private Sub CleanCacheAll()

   Dim I As Integer

   DictIdListIdx = 0
   Client.Server.CreateDictIdList 0, 0, 0
   Do While Client.Server.DictIdGetNext(DictIdList(DictIdListIdx))
      If DictIdListIdx < UBound(DictIdList) Then
         DictIdListIdx = DictIdListIdx + 1
      End If
   Loop
   
   For I = 0 To DictCacheNumber - 1
      If Not IsDictIdInDictIdList(DictCache(I).DictId) Then
         DeleteDictFromCache I
      End If
   Next I
End Sub
Private Function IsDictIdInDictIdList(DictId As Long) As Boolean

   Dim H As Long
   Dim L As Long
   Dim I As Long
   
   H = DictIdListIdx - 1
   L = 0
   Do
      I = (H + L) / 2
      If H - L > 1 Then
         If DictIdList(I) > DictId Then
            H = I
         ElseIf DictIdList(I) < DictId Then
            L = I
         Else
            IsDictIdInDictIdList = True
            Exit Function
         End If
      Else
         If DictIdList(L) = DictId Or DictIdList(H) = DictId Then
            IsDictIdInDictIdList = True
            Exit Function
         Else
            L = H
         End If
      End If
   Loop While H <> L

   IsDictIdInDictIdList = False
End Function
Private Function FillCacheFromFilter(ByVal OrgId As Long, ByRef TooMany As Boolean) As Double
 
   Dim Dict As clsDict
   Dim MaxTimeStamp As Double
   Dim Flt As clsFilter
   Dim FirstUpdate As Boolean
   Dim SQLSearch As String
   Dim NumberDictInSearchResult As Integer

   TooMany = False
   For Each Flt In FilterCol
      If Flt.OrgId <> 0 Then
         If DictCacheTimeStampOrg(Flt.OrgId) < 0 Then
            FirstUpdate = True
            Client.Server.CreateDictList Flt.OrgId, 0, 0, Recorded, Recorded, 0
            Do While Client.Server.DictListGetNext(Dict)
               If MaxTimeStamp < Dict.TimeStamp Then
                  MaxTimeStamp = Dict.TimeStamp
               End If
               UpdateOneDictInCache Dict, FirstUpdate
               FirstUpdate = False
            Loop
            DictCacheTimeStampOrg(Flt.OrgId) = MaxTimeStamp
         Else
            MaxTimeStamp = DictCacheTimeStampOrg(Flt.OrgId)
         End If
      ElseIf Flt.AuthorId <> 0 Then
            FirstUpdate = True
            Client.Server.CreateDictList 0, Flt.AuthorId, 0, Flt.StatusStart, Flt.StatusEnd, 0
            Do While Client.Server.DictListGetNext(Dict)
               If MaxTimeStamp < Dict.TimeStamp Then
                  MaxTimeStamp = Dict.TimeStamp
               End If
               UpdateOneDictInCache Dict, FirstUpdate
               FirstUpdate = False
            Loop
      ElseIf Flt.TranscriberId <> 0 Then
         FirstUpdate = True
         Client.Server.CreateDictList 0, 0, Flt.TranscriberId, Flt.StatusStart, Flt.StatusEnd, 0
         Do While Client.Server.DictListGetNext(Dict)
            If MaxTimeStamp < Dict.TimeStamp Then
               MaxTimeStamp = Dict.TimeStamp
            End If
            UpdateOneDictInCache Dict, FirstUpdate
            FirstUpdate = False
         Loop
      ElseIf OrgId = 30005 Then
         FirstUpdate = True
         Client.Server.CreateDictList 0, 0, 0, mCurrPatientFilter.StatusStart, mCurrPatientFilter.StatusEnd, 0, mCurrPatientFilter.Pat.PatId, mCurrPatientFilter.Pat.PatId2
         Do While Client.Server.DictListGetNext(Dict)
            If MaxTimeStamp < Dict.TimeStamp Then
               MaxTimeStamp = Dict.TimeStamp
            End If
            UpdateOneDictInCache Dict, FirstUpdate
            FirstUpdate = False
         Loop
      Else
         Dim T As Variant
         T = Timer
         FirstUpdate = True
         SQLSearch = CreateSQLFromFilter(Flt)
         Client.Server.CreateDictListFromSQL SQLSearch
         NumberDictInSearchResult = 0
         Do While Client.Server.DictListGetNext(Dict)
            If MaxTimeStamp < Dict.TimeStamp Then
               MaxTimeStamp = Dict.TimeStamp
            End If
            NumberDictInSearchResult = NumberDictInSearchResult + 1
            UpdateOneDictInCache Dict, FirstUpdate
            FirstUpdate = False
         Loop
         If NumberDictInSearchResult >= MaxNumberDictInSearchResult Then
            TooMany = True
         End If
         Debug.Print "Tid: " & Format$(Timer - T, "0.00")
      End If
   Next Flt
   FillCacheFromFilter = MaxTimeStamp
End Function
Private Function CreateSQLFromFilter(Flt As clsFilter) As String

   Dim SelectClause As String
   Dim WhereClause As String
   Dim OrderClause As String
   Dim Txt As String
   Dim TxtWord As String
   
   SelectClause = "Select top " & CStr(MaxNumberDictInSearchResult) & " * from v_Dictation_List"
   
   If Len(RemoveDelimiters(Flt.Pat.PatId)) > 0 Then
      Client.Server.AddString WhereClause, "and", "PatId like '%" & RemoveDelimiters(Flt.Pat.PatId) & "%'"
   End If
   If Len(RemoveDelimiters(Flt.Pat.PatName)) > 0 Then
      Client.Server.AddString WhereClause, "and", "PatName like '%" & RemoveDelimiters(Flt.Pat.PatName) & "%'"
   End If
   
   If Flt.RecDateUsed Then
      Client.Server.AddString WhereClause, "and", "Created>=" & Client.Server.SQLDateDelimiter & Format$(Flt.RecStartDate, Client.Server.SQLDateStringFormat) & Client.Server.SQLDateDelimiter
      Client.Server.AddString WhereClause, "and", "Created<=" & Client.Server.SQLDateDelimiter & Format$(Flt.RecEndDate, Client.Server.SQLDateStringFormat) & Client.Server.SQLDateDelimiter
   End If
   
   If Flt.TranscribedDateUsed Then
      Client.Server.AddString WhereClause, "and", "TranscribedDate>=" & Client.Server.SQLDateDelimiter & Format$(Flt.TranscribedStartDate, Client.Server.SQLDateStringFormat) & Client.Server.SQLDateDelimiter
      Client.Server.AddString WhereClause, "and", "TranscribedDate<=" & Client.Server.SQLDateDelimiter & Format$(Flt.TranscribedEndDate, Client.Server.SQLDateStringFormat) & Client.Server.SQLDateDelimiter
   End If
   
   If Flt.DictTypeId >= 0 Then
      Client.Server.AddString WhereClause, "and", "DictTypeId=" & CStr(Flt.DictTypeId)
   End If
   If Flt.PriorityId >= 0 Then
      Client.Server.AddString WhereClause, "and", "PriorityId=" & CStr(Flt.PriorityId)
   End If
   If Len(RemoveDelimiters(Flt.AuthorName)) > 0 Then
      Client.Server.AddString WhereClause, "and", "AuthorShortName like '%" & RemoveDelimiters(Flt.AuthorName) & "%'"
   End If
   If Len(RemoveDelimiters(Flt.TranscriberName)) > 0 Then
      Client.Server.AddString WhereClause, "and", "TranscriberShortName like '%" & RemoveDelimiters(Flt.TranscriberName) & "%'"
   End If
   Txt = RemoveDelimiters(Flt.Txt)
   Do
      TxtWord = GetNextWordFromTxt(Txt)
      If Len(TxtWord) > 0 Then
         Client.Server.AddString WhereClause, "and", "Txt like '%" & TxtWord & "%'"
      End If
   Loop Until Len(TxtWord) = 0
   
   OrderClause = "DictId"
   
   CreateSQLFromFilter = Client.Server.BuildSQL(SelectClause, WhereClause, "", OrderClause)
End Function
Private Function RemoveDelimiters(S As String) As String

   Dim T As String
   
   T = Trim$(S)
   T = Replace(T, "%", "")
   T = Replace(T, "'", "")
   T = Replace(T, "/", "")
   T = Replace(T, "\", "")
   T = Replace(T, "-", "")
   T = Replace(T, "_", "")
   T = Replace(T, ".", "")
   T = Replace(T, ",", "")
   T = Replace(T, "?", "")
   T = Replace(T, ":", "")
   T = Replace(T, ";", "")
   T = Replace(T, "&", "")
   T = Replace(T, """", "")
   T = Replace(T, "!", "")
   RemoveDelimiters = T
End Function

Private Function GetNextWordFromTxt(ByRef Txt As String) As String

   Dim Pos As Integer
   
   Txt = Trim$(Txt)
   Pos = InStr(Txt, " ")
   If Pos > 0 Then
      GetNextWordFromTxt = Trim$(Left$(Txt, Pos - 1))
      Txt = Trim$(mId$(Txt, Pos + 1))
   Else
      GetNextWordFromTxt = Txt
      Txt = ""
   End If
End Function
Private Sub BuildFilterForOrgId(OrgId As Long)

   AddFilter OrgId, 0, 0, Recorded, Recorded
   BuildFilterForChildren OrgId
End Sub
Private Sub BuildFilterForChildren(OrgParentId As Long)

   Dim I As Integer
   Dim Org As clsOrg
   
   For I = 0 To Client.OrgMgr.Count - 1
      Client.OrgMgr.GetSortedOrg Org, I
      If Org.OrgParent = OrgParentId Then
         AddFilter Org.OrgId, 0, 0, Recorded, Recorded
         BuildFilterForChildren Org.OrgId
      End If
   Next I
End Sub
Private Sub AddFilter(OrgId As Long, AuthorId As Long, TranscriberId As Long, StatusStart As Integer, StatusEnd As Integer)

   Dim Filter As New clsFilter
   
   Filter.OrgId = OrgId
   Filter.AuthorId = AuthorId
   Filter.TranscriberId = TranscriberId
   Filter.StatusStart = StatusStart
   Filter.StatusEnd = StatusEnd
   
   FilterCol.Add Filter
End Sub
Private Function FilterDictation(Dict As clsDict) As Boolean

   Dim Flt As clsFilter
   Dim FilteredByPatId As Boolean
   Dim FilteredByRecDate As Boolean
   Dim FilteredByTranscribedDate As Boolean

   If Not Client.OrgMgr.CheckUserRoleHidden(Dict) Then
      
      For Each Flt In FilterCol
      
         FilteredByPatId = False
         If Len(Flt.Pat.PatId) = 0 And Len(Flt.Pat.PatId2) = 0 Then
            FilteredByPatId = True
         ElseIf Len(Flt.Pat.PatId2) = 0 Then
            FilteredByPatId = InStr(UCase$(Dict.Pat.PatId), UCase$(Flt.Pat.PatId)) > 0
         ElseIf Len(Flt.Pat.PatId) = 0 Then
            FilteredByPatId = InStr(UCase$(Dict.Pat.PatId), UCase$(Flt.Pat.PatId2)) > 0
         Else
            FilteredByPatId = InStr(UCase$(Dict.Pat.PatId), UCase$(Flt.Pat.PatId)) > 0 Or InStr(UCase$(Dict.Pat.PatId), UCase$(Flt.Pat.PatId2)) > 0
         End If
         
         If Flt.RecDateUsed Then
            FilteredByRecDate = Dict.Created >= Flt.RecStartDate And Dict.Created <= Flt.RecEndDate
         Else
            FilteredByRecDate = True
         End If
      
         If Flt.TranscribedDateUsed Then
            FilteredByTranscribedDate = Dict.TranscribedDate >= Flt.TranscribedStartDate And Dict.TranscribedDate <= Flt.TranscribedEndDate
         Else
            FilteredByTranscribedDate = True
         End If
      
      
         If (Flt.OrgId = Dict.OrgId Or Flt.OrgId = 0) And _
            (Flt.AuthorId = Dict.AuthorId Or Flt.AuthorId = 0) And _
            (Flt.TranscriberId = Dict.TranscriberId Or Flt.TranscriberId = 0) And _
            (Flt.StatusStart <= Dict.StatusId) And _
            (Flt.StatusEnd >= Dict.StatusId) And _
            (Flt.DictTypeId = Dict.DictTypeId Or Flt.DictTypeId < 0) And _
            (Flt.PriorityId = Dict.PriorityId Or Flt.PriorityId < 0) And _
            FilteredByPatId And _
            FilteredByRecDate And _
            FilteredByTranscribedDate And _
            (InStr(UCase$(Dict.Pat.PatName), UCase$(Flt.Pat.PatName)) > 0 Or Len(Flt.Pat.PatName) = 0) And _
            (InStr(UCase$(Dict.AuthorShortName), UCase$(Flt.AuthorName)) > 0) And _
            (InStr(UCase$(Dict.TranscriberShortName), UCase$(Flt.TranscriberName)) > 0 Or Len(Flt.TranscriberName) = 0) And _
            FilterByTxt(Dict.Txt, Flt.Txt) Then
            
            FilterDictation = True
            Exit For
         End If
      Next Flt
   'Else
   '   MsgBox "0"
   End If
End Function
Private Function FilterByTxt(ByVal Dtxt As String, ByVal Ftxt As String) As Boolean

   Dim Word As String
   
   Dtxt = UCase$(Dtxt)
   Ftxt = UCase$(Ftxt)
   Do While Len(Ftxt) > 0
      Word = GetWordFromString(Ftxt)
      If Len(Word) > 0 Then
         If InStr(Dtxt, Word) = 0 Then
            FilterByTxt = False
            Exit Function
         Else
            Debug.Print Dtxt
         End If
      End If
   Loop
   FilterByTxt = True
End Function
Private Function GetWordFromString(ByRef S As String) As String

   Dim Pos As Integer
   
   Pos = InStr(S, " ")
   If Pos > 0 Then
      GetWordFromString = Trim$(Left$(S, Pos - 1))
      S = Trim$(mId$(S, Pos + 1))
   Else
      GetWordFromString = Trim$(S)
      S = ""
   End If
End Function
Private Sub UpdateOneDictInCache(ByVal Dict As clsDict, FirstUpdate As Boolean)

   Dim I As Integer
   Static Idx As Integer
   
   Dict.StatusText = Client.Texts.Txt(1250100 + Dict.StatusId, Dict.StatusText)
   If FirstUpdate Then Idx = 0
   Do While Idx < DictCacheNumber
      If DictCache(Idx).DictId = Dict.DictId Then
         Set DictCache(Idx) = Dict
         If Not Client.OrgMgr.CheckUserRole(Dict.OrgId, RTList) Then
            DeleteDictFromCache Idx
         End If
         Exit Sub
      ElseIf DictCache(Idx).DictId > Dict.DictId Then
         If Client.OrgMgr.CheckUserRole(Dict.OrgId, RTList) Then
            For I = DictCacheNumber To Idx + 1 Step -1
               Set DictCache(I) = DictCache(I - 1)
            Next I
            DictCacheNumber = DictCacheNumber + 1
            Set DictCache(Idx) = Dict
         End If
         Exit Sub
      End If
      Idx = Idx + 1
   Loop
   If Client.OrgMgr.CheckUserRole(Dict.OrgId, RTList) Then
      Set DictCache(DictCacheNumber) = Dict
      DictCacheNumber = DictCacheNumber + 1
   End If
End Sub
Private Sub DeleteDictFromCache(Index As Integer)

   Dim I As Integer

   For I = Index To DictCacheNumber - 2
      Set DictCache(I) = DictCache(I + 1)
   Next I
   DictCacheNumber = DictCacheNumber - 1
End Sub
Public Function CheckOut(ByRef Dict As clsDict, ByVal DictId As Long, ByVal WithSound As Boolean) As Integer

   Dim Res As Integer
   Dim TempfileName As String
   
   If WithSound Then
      TempfileName = CreateTempFileName("dss")
   End If
   Res = Client.Server.CheckOutDict(Dict, DictId, WithSound, TempfileName)
   If Res = 0 Then
      Client.DictAuditMgr.Insert DictId, AuditType_CheckOut, Dict.StatusId
      Dict.SoundReadOnly = Dict.AuthorId <> Client.User.UserId Or Dict.StatusId > 30
      If Not Dict.SoundReadOnly Then
         Dict.TextReadOnly = False
      Else
         Dict.TextReadOnly = Not (Client.OrgMgr.CheckUserRole(Dict.OrgId, RTTextEdit) And Dict.StatusId < SoundDeleted)
      End If
   Else
      If WithSound Then
         KillFileIgnoreError TempfileName
      End If
      MsgBox Client.Texts.Txt(1190101, "Diktatet är upptaget")
   End If
   CheckOut = Res
End Function
Public Function CheckIn(ByRef Dict As clsDict, ByVal DiscardChanges As Boolean) As Integer

   Dim Res As Integer
   
   If Dict.InfoDirty Then
      Client.DictAuditMgr.Insert Dict.DictId, AuditType_InfoChanged, Dict.StatusId
   End If
   If Dict.SoundDirty Then
      Client.DictAuditMgr.Insert Dict.DictId, AuditType_SoundChanged, Dict.StatusId
   End If
   Res = Client.Server.CheckInDict(Dict, DiscardChanges)
   If Res = 0 Then
      Client.DictAuditMgr.Insert Dict.DictId, AuditType_CheckIn, Dict.StatusId
   Else
      MsgBox Client.Texts.Txt(1190102, "Diktatet kunde inte sparas")
   End If
   If Len(Dict.LocalFilename) > 0 Then
      KillFileIgnoreError Dict.LocalFilename
      Dict.LocalFilename = ""
   End If
   CheckIn = Res
End Function
Public Function CreateNew(ByRef Dict As clsDict) As Integer

   Dim Res As Integer
   Dim Idx As Integer
   Dim TempfileName As String
      
   Dict.Created = Now
   If Len(Dict.LocalFilename) = 0 Then
      TempfileName = CreateTempFileName("dss")
      Dict.LocalFilename = TempfileName
   End If
   Dict.AuthorId = Client.User.UserId
   Dict.AuthorLongName = Client.User.LongName
   Dict.AuthorShortName = Client.User.ShortName
   
   Dict.PriorityId = -1
   Dict.DictTypeId = -1
   CreateNew = 0
End Function
Public Function CheckInNew(ByRef Dict As clsDict) As Integer

   Dim Res As Integer
   
   Res = Client.Server.CheckInNewDict(Dict)
   If Res = 0 Then
      Client.LoggMgr.Insert 1320112, LoggLevel_DictInfo, Dict.DictId, Dict.LoggData
      Client.DictAuditMgr.Insert Dict.DictId, AuditType_CheckInNew, Dict.StatusId
   Else
      Client.LoggMgr.Insert 1320113, LoggLevel_DictFailure, Dict.DictId, Dict.LoggData
      MsgBox Client.Texts.Txt(1190102, "Diktatet kunde inte sparas")
   End If
   
   If Len(Dict.LocalFilename) Then
      KillFileIgnoreError Dict.LocalFilename
      Dict.LocalFilename = ""
   End If
   CheckInNew = Res
End Function

Public Function IsThereDictations(OrgId As Long) As Boolean

   Dim Cnt As Integer
   Dim D As clsDict
   Dim TooMany As Boolean
   
   CreateList OrgId, 0, TooMany
   Do While ListNextItem(D)
      Cnt = Cnt + 1
   Loop
   IsThereDictations = (Cnt > 0)
End Function

Private Sub Class_Initialize()

   Dim I As Integer
   
   For I = LBound(DictCacheTimeStampOrg) To UBound(DictCacheTimeStampOrg)
      DictCacheTimeStampOrg(I) = -1
   Next I
End Sub

Public Sub SaveTempDictationInfo(D As clsDict, Mode As TempDictInfoTypeEnum)

   Dim Fn As String
   Dim Settings As New clsSettings
   Dim Section As String
   
   Fn = CreateTempPath() & "CtRecovery.dat"
   Settings.IniFilePath = Fn
   
   Section = CStr(Client.User.UserId)
   If Mode <> tdiUpdateInfo Then
      Settings.WriteLong Section, "Mode", CLng(Mode)
   End If
   Settings.WriteLong Section, "DictId", D.DictId
   Settings.WriteLong Section, "Status", D.StatusId
   Settings.WriteLong Section, "OrgId", D.OrgId
   Settings.WriteString Section, "PatId", D.Pat.PatId
   Settings.WriteBool Section, "NoPat", D.NoPatient
   Settings.WriteString Section, "PatName", D.Pat.PatName
   Settings.WriteLong Section, "DictType", D.DictTypeId
   Settings.WriteLong Section, "Priority", D.PriorityId
   Settings.WriteString Section, "Txt", D.Txt
   Settings.WriteString Section, "Lfn", D.LocalFilename
End Sub
Public Sub EmptyTempDictationInfo()

   Dim EmptyDict As New clsDict
   
   SaveTempDictationInfo EmptyDict, tdiEmpty
End Sub
Public Function RestoreTempDictationInfo(D As clsDict) As TempDictInfoTypeEnum

   Dim Fn As String
   Dim Settings As New clsSettings
   Dim Section As String
   Dim Mode As TempDictInfoTypeEnum
   
   Fn = CreateTempPath() & "CtRecovery.dat"
   Settings.IniFilePath = Fn
   
   Section = CStr(Client.User.UserId)
   Mode = Settings.GetLong(Section, "Mode", tdiEmpty)
   RestoreTempDictationInfo = Mode
   If Mode <> tdiEmpty Then
      Set D = New clsDict
      D.DictId = Settings.GetLong(Section, "DictId", 0)
      D.StatusId = Settings.GetLong(Section, "Status", 0)
      D.OrgId = Settings.GetLong(Section, "OrgId", 0)
      D.Pat.PatId = Settings.GetString(Section, "PatId", "")
      D.NoPatient = Settings.GetBool(Section, "NoPat", False)
      D.Pat.PatName = Settings.GetString(Section, "PatName", "")
      D.DictTypeId = Settings.GetLong(Section, "DictType", 0)
      D.PriorityId = Settings.GetLong(Section, "Priority", 0)
      D.Txt = Settings.GetString(Section, "Txt", "")
      D.LocalFilename = Settings.GetString(Section, "Lfn", "")
      
      D.AuthorId = Client.User.UserId
   End If
End Function

