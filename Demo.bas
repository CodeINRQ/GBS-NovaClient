Attribute VB_Name = "modDemo"
Option Explicit

Sub FillDemoDictation(NumberOfDictations As Long)

   Dim I As Long
   Dim D As clsDict
   
   For I = 1 To NumberOfDictations
      DoEvents
      Set D = New clsDict
      
      D.Pat.PatId = Client.Custom.RndPatId()
      D.Pat.PatName = Client.Custom.RndPatName()
      D.StatusId = RndStatus()
      If D.StatusId < Transcribed Then
         D.Created = rndDate(6)
      Else
         D.Created = rndDate(365)
      End If
      D.SoundLength = Int(Rnd * 60 * 40) + 20
      D.OrgId = RndOrg()
      D.DictTypeId = RndDictType
      D.AuthorId = RndUser(D.OrgId, "A")
      If D.StatusId >= Transcribed Then
         D.TranscriberId = RndUser(D.OrgId, "T")
         D.TranscribedDate = DateAdd("h", Int(Rnd * 5 * 24), D.Created)
      End If
      If Rnd < 0.1 Then
         D.Changed = DateAdd("h", 5, D.Created)
      End If
      
      Dim Prio As clsPriority
      Client.PriorityMgr.GetFromId Prio, RndPrio()
      D.PriorityId = Prio.PriorityId
      D.ExpiryDate = DateAdd("d", Prio.Days, D.Created)
      Set Prio = Nothing
      
      If D.StatusId < SoundDeleted Then
         D.LocalFilename = RndDictFile()
      End If
      Client.Server.CheckInNewDemo D
      Client.DictAuditMgr.Insert D.DictId, 12, D.StatusId
      Set D = Nothing
   Next I
   
   'FillHistoryDemo
End Sub

Sub FillHistoryDemo(Days As Integer, MaxNumberPerDay As Integer, MinNumberPerDay As Integer)

   Dim Rs As New ADODB.Recordset
   Dim Cr As Integer
   Dim I As Integer
   Dim r As Integer
   Dim Ex As Integer
   Dim DataB As ADODB.Connection
   
   Set DataB = Client.Server.Connection
   Rs.Open "Select * from History", DataB, adOpenDynamic, adLockPessimistic
   For Cr = 1 To Days
      r = Int(Rnd * MaxNumberPerDay - MinNumberPerDay) + MinNumberPerDay
      For I = 1 To r
         Rs.AddNew
         Rs("DictId") = CLng(Cr) * CLng(1000) + CLng(I)
         Rs("Created") = DateAdd("d", -Cr, Now)
         Rs("OrgId") = RndOrg()
         Rs("DictTypeId") = RndDictType()
         
         Dim Prio As clsPriority
         Client.PriorityMgr.GetFromId Prio, RndPrio()
         Rs("PriorityId") = Prio.PriorityId
         Rs("ExpiryDate") = DateAdd("d", Prio.Days, Rs("created"))
         Set Prio = Nothing
         
         Rs("AuthorId") = RndUser(Rs("OrgId"), "A")
         Rs("TranscriberId") = RndUser(Rs("OrgId"), "T")
         Rs("TranscriberOrgId") = RndOrg
         Rs("TranscribedDate") = DateAdd("h", Int(Rnd * 5 * 24), Rs("Created"))
         Rs("SoundLenSec") = Int(Rnd * 240) + 60
         
         Rs.Update
      Next I
   Next Cr
   Rs.Close
   Set Rs = Nothing
End Sub
Function RndDictFile() As String

   Dim TFn As String
   Dim I As Integer
   
   TFn = CreateTempFileName("")
   FileCopy App.Path & "\DemoDict\" & CStr(Int((Rnd * 4) + 1)) & ".dss", TFn
   RndDictFile = TFn
End Function
Function RndOrg() As Long

   Dim Org As New clsOrg
   Dim OrgIdx As Integer
   
   Do While True
      OrgIdx = Int(Rnd * Client.OrgMgr.Count)
      Client.OrgMgr.GetSortedOrg Org, OrgIdx
      If Org.DictContainer Then
         RndOrg = Org.OrgId
         Exit Function
      End If
   Loop
End Function
Function RndUser(OrgId As Long, Priv As String) As Long

   Dim Rs As New ADODB.Recordset
   Dim I As Integer
   Dim IRnd As Integer
   Dim UserId As Long
   Dim Roles As String
   Dim DataB As ADODB.Connection
   
   Set DataB = Client.Server.Connection
   Do While True
      Rs.Open "spUserRole", DataB, adOpenDynamic, adLockReadOnly
      IRnd = Int(Rnd * 20)
      For I = 0 To IRnd
         If Rs.EOF Then
            Rs.MoveFirst
         Else
            Rs.MoveNext
         End If
      Next I
      If Rs.EOF Or Rs.BOF Then
         Rs.MoveFirst
      End If
      If IsNull(Rs("UserId")) Then
         Rs.MoveFirst
      End If
      UserId = Rs("UserId")
      Roles = Rs("Roles")
      Rs.Close
      
      If InStr(Roles, Priv) > 0 And UserId <> Client.User.UserId Then
         RndUser = UserId
         Exit Function
      End If
   Loop
End Function
Function RndDictType() As Long

   Dim D As New clsDictType
   Dim Idx As Integer
   
   Idx = Int(Rnd * Client.DictTypeMgr.Count)
   RndDictType = Client.DictTypeMgr.IdFromIndex(Idx)
End Function
Function RndStatus() As Long

   Dim I As Integer
   
   I = Int(Rnd * 16)
   Select Case I
      Case 0: RndStatus = Recorded
      Case 1: RndStatus = Recorded
      Case 2: RndStatus = Recorded
      Case 3: RndStatus = Recorded
      Case 4: RndStatus = Recorded
      Case 5: RndStatus = Recorded
      Case 6: RndStatus = Recorded
      Case 7: RndStatus = Transcribed
      Case 8: RndStatus = Transcribed
      Case 9: RndStatus = Transcribed
      Case 10: RndStatus = Transcribed
      Case 11: RndStatus = Transcribed
      Case 12: RndStatus = Transcribed
      Case 13: RndStatus = Transcribed
      Case 14: RndStatus = Transcribed
      Case 15: RndStatus = SoundDeleted
   End Select
End Function
Function RndPrio() As Long

   Dim D As New clsPriority
   Dim Idx As Integer
   
   Idx = Int(Rnd * Client.PriorityMgr.Count)
   RndPrio = Client.PriorityMgr.IdFromIndex(Idx)
End Function

Function rndDate(DaysOld As Integer) As Date

   Dim I As Integer
   
   I = Int(Rnd * DaysOld)
   rndDate = DateAdd("n", -Int(Rnd * 60 * 24), DateAdd("d", -I, Now))
End Function
