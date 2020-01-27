VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrgMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOrgArr() As clsOrg
Private mOrgCount As Integer
Private mOrgTreeOrder() As Integer
Private mShowAll As Boolean

Public Sub Init(ShowAll As Boolean)
   
   If ShowAll <> mShowAll Then
      mShowAll = ShowAll
      mOrgCount = 0
   End If
   If mOrgCount = 0 Then
      Client.Server.CreateOrgList
      Do While Client.Server.OrgListGetNext(mOrgArr(mOrgCount))
         mOrgCount = mOrgCount + 1
         If mOrgCount = UBound(mOrgArr) Then
            ReDim Preserve mOrgArr(mOrgCount + 50)
            ReDim Preserve mOrgTreeOrder(mOrgCount + 50)
         End If
      Loop
      
      SetTreeInfo 0, 0, 0
      SortForShowInTree
   End If
End Sub
Public Sub SaveOrg(Org As clsOrg)

   On Error GoTo SaveOrg_Err
   If Org.OrgParent <> 0 Then
      Client.Server.SaveOrg Org
      Client.LoggMgr.Insert 1320111, LoggLevel_SysAdmin, 0, Org.LoggData
   End If
   mOrgCount = 0
   Init mShowAll
   Exit Sub
   
SaveOrg_Err:
   ErrorHandle "1280101", Err, 1280101, "Enheten kan inte sparas", False
   Resume Next
End Sub
Private Sub SetTreeInfo(ByVal Level As Integer, ByVal NodeId As Long, ByVal CurrentShowParent As Long)

   Dim Idx As Integer
   Dim NewShowParent As Long
   
   For Idx = 0 To mOrgCount - 1
      With mOrgArr(Idx)
         If .OrgParent = NodeId Then
            .TreeLevel = Level
            .ShowParent = CurrentShowParent
            If .ShowInTree Or mShowAll Then
               NewShowParent = .OrgId
            Else
               NewShowParent = CurrentShowParent
            End If
            SetTreeInfo Level + 1, mOrgArr(Idx).OrgId, NewShowParent
         End If
      End With
   Next Idx
End Sub
Private Sub SortForShowInTree()

   Dim I As Integer
   Dim J As Integer
   Dim T As Integer
   
   For I = 0 To mOrgCount - 1
      mOrgTreeOrder(I) = I
   Next I
   
   For I = 0 To mOrgCount - 2
      For J = I + 1 To mOrgCount - 1
         If (mOrgArr(mOrgTreeOrder(I)).TreeLevel > mOrgArr(mOrgTreeOrder(J)).TreeLevel) Or _
            ((mOrgArr(mOrgTreeOrder(I)).TreeLevel = mOrgArr(mOrgTreeOrder(J)).TreeLevel) And _
             (UCase$(mOrgArr(mOrgTreeOrder(I)).OrgText) > UCase$(mOrgArr(mOrgTreeOrder(J)).OrgText))) Then
            T = mOrgTreeOrder(J)
            mOrgTreeOrder(J) = mOrgTreeOrder(I)
            mOrgTreeOrder(I) = T
         End If
      Next J
   Next I
End Sub
Public Property Get Count() As Integer

   Count = mOrgCount
End Property
Public Sub GetSortedOrg(ByRef Org As clsOrg, Idx As Integer)

   Set Org = mOrgArr(mOrgTreeOrder(Idx))
End Sub
Public Sub GetOrgFromId(ByRef Org As clsOrg, OrgId As Long)

   Dim I As Integer
   
   For I = 0 To mOrgCount - 1
      If mOrgArr(I).OrgId = OrgId Then
         Set Org = mOrgArr(I)
         Exit For
      End If
   Next
End Sub
Public Function TextFromId(OrgId As Long) As String

   Dim I As Integer
   
   For I = 0 To mOrgCount - 1
      If mOrgArr(I).OrgId = OrgId Then
         TextFromId = mOrgArr(I).OrgText
         Exit For
      End If
   Next
End Function
Public Function CheckUserRole(ByVal OrgId As Long, ByVal Role As String) As Boolean

   Dim I As Integer
   Dim Org As clsOrg
   
   If OrgId > 0 Then
      GetOrgFromId Org, OrgId
      If Not Org Is Nothing Then
         If InStr(Org.Roles.Roles, Role) > 0 Then
            CheckUserRole = True
         End If
      End If
   Else
      For I = 0 To mOrgCount - 1
         If InStr(mOrgArr(I).Roles.Roles, Role) > 0 Then
            CheckUserRole = True
            Exit For
         End If
      Next I
   End If
End Function
Public Function CheckUserAllowListening(ByVal OrgId As Long) As Boolean

   Dim I As Integer
   Dim Org As clsOrg
   
   If OrgId > 0 Then
      GetOrgFromId Org, OrgId
      If Not Org Is Nothing Then
         If Len(Org.Roles.Roles) > 0 Then
            CheckUserAllowListening = True
         End If
      End If
   Else
      CheckUserAllowListening = True
   End If
End Function
Public Function CheckUserRoleHidden(D As clsDict) As Boolean

   Dim I As Integer
   Dim Org As clsOrg
   Dim FirstTimeToShow As Date
   
   If D.StatusId < Transcribed And D.AuthorId <> Client.User.UserId And D.TranscriberId <> Client.User.UserId Then
      GetOrgFromId Org, D.OrgId
      If Not Org Is Nothing Then
         If Org.Roles.Delayed Then
            FirstTimeToShow = DateAdd("h", Org.Roles.DelayedHours, D.ExpiryDate)
            CheckUserRoleHidden = FirstTimeToShow > Now
         End If
      End If
   End If
End Function
Private Sub Class_Initialize()

   ReDim mOrgArr(50)
   ReDim mOrgTreeOrder(50)
End Sub
