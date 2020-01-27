VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StatCache(10000) As clsStat
Private StatCacheNumber As Integer
Private StatCacheIndex As Integer

Public Sub CreateList(ByVal OrgId As Long)

   StatCacheNumber = 0
   StatCacheIndex = 0
   Select Case OrgId
      Case Is < 30000
         AddStatForOrgId OrgId
   End Select
End Sub
Private Sub AddStatForOrgId(OrgId As Long)

   AddStat OrgId
   AddStatForChildren OrgId
End Sub
Private Sub AddStatForChildren(OrgParentId As Long)

   Dim I As Integer
   Dim Org As clsOrg
   
   For I = 0 To Client.OrgMgr.Count - 1
      Client.OrgMgr.GetSortedOrg Org, I
      If Org.OrgParent = OrgParentId Then
         AddStat Org.OrgId
         AddStatForChildren Org.OrgId
      End If
   Next I
End Sub
Private Sub AddStat(OrgId As Long)

   Dim Stat As clsStat
   Dim Org As clsOrg
   
   Client.OrgMgr.GetOrgFromId Org, OrgId
   If Org.DictContainer Then
      Client.Server.CreateStat OrgId, Stat, Recorded, BeingTrancribed, -4, 0, 1, 2, 3, 5, 10
      Set StatCache(StatCacheNumber) = Stat
      StatCacheNumber = StatCacheNumber + 1
   End If
End Sub

Public Function ListNextItem(ByRef Stat As clsStat) As Boolean

   Dim Hit As Boolean

   Do While StatCacheIndex < StatCacheNumber
      Hit = True
      Set Stat = StatCache(StatCacheIndex)
      StatCacheIndex = StatCacheIndex + 1
      If Hit Then
         ListNextItem = True
         Exit Function
      End If
   Loop
   ListNextItem = False
End Function


