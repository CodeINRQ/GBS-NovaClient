VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoryMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum HistTypeEnum
   htPrio = 0
   htDictType = 1
   htOrg = 2
   htTranscriber = 3
   htTranscriberOrg = 4
   htDaysFromCreated = 5
   htDaysFromExpiry = 6
   htAuthor = 7
   htminutesFromCreated = 8
   htminutesFromExpiry = 9
End Enum

Private mHistyear As Integer
Private mHistType As HistTypeEnum

Private Cache(10000) As clsHistory
Private CacheNumber As Integer
Private CacheIndex As Integer

Public Function CreateList(ByVal OrgId As Long, HistType As HistTypeEnum, HistYear As Integer)

   CacheNumber = 0
   CacheIndex = 0
   mHistyear = HistYear
   mHistType = HistType
        
   Select Case OrgId
      Case Is < 30000
         AddHistForOrgId OrgId
   End Select
End Function
Private Sub AddHistForOrgId(OrgId As Long)

   AddHist OrgId
   AddHistForChildren OrgId
End Sub
Private Sub AddHistForChildren(OrgParentId As Long)

   Dim I As Integer
   Dim Org As clsOrg
   
   For I = 0 To Client.OrgMgr.Count - 1
      Client.OrgMgr.GetSortedOrg Org, I
      If Org.OrgParent = OrgParentId Then
         AddHist Org.OrgId
         AddHistForChildren Org.OrgId
      End If
   Next I
End Sub
Private Sub AddHist(OrgId As Long)

   Dim I As Integer
   Dim Mth As Integer
   Dim Hist As clsHistory
   Dim Hit As Boolean
   
   Client.Server.CreateHist OrgId, mHistyear, mHistType
   Do While Client.Server.HistGetNext(Hist)
      Hit = False
      For I = 0 To CacheNumber - 1
         If Hist.Rowid = Cache(I).Rowid Then
            Hit = True
            Exit For
         End If
      Next I
      If Hit Then
         For Mth = 1 To 12
            Cache(I).Number(Mth) = Cache(I).Number(Mth) + Hist.Number(Mth)
            Cache(I).SoundLenSec(Mth) = Cache(I).SoundLenSec(Mth) + Hist.SoundLenSec(Mth)
         Next Mth
      Else
         Set Cache(CacheNumber) = Hist
         CacheNumber = CacheNumber + 1
      End If
   Loop
End Sub

Public Function ListNextItem(ByRef Hist As clsHistory) As Boolean

   Dim Hit As Boolean

   Do While CacheIndex < CacheNumber
      Hit = True
      Set Hist = Cache(CacheIndex)
      CacheIndex = CacheIndex + 1
      If Hit Then
         ListNextItem = True
         Exit Function
      End If
   Loop
   ListNextItem = False
End Function



