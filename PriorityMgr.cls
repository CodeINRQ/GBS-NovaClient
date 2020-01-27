VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPriorityMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPriorityArr() As clsPriority
Private mPriorityCount As Integer

Private mOrgPriorityArr() As clsOrgPriority
Private mOrgPriorityCount As Integer

Public Sub Init()

   mPriorityCount = 0
   Client.Server.CreatePriorityList
   Do While Client.Server.PriorityListGetNext(mPriorityArr(mPriorityCount))
      mPriorityCount = mPriorityCount + 1
      If mPriorityCount = UBound(mPriorityArr) Then
         ReDim Preserve mPriorityArr(mPriorityCount + 20)
      End If
   Loop
   
   mOrgPriorityCount = 0
   Client.Server.CreateOrgPriorityList
   Do While Client.Server.OrgPriorityListGetNext(mOrgPriorityArr(mOrgPriorityCount))
      mOrgPriorityCount = mOrgPriorityCount + 1
      If mOrgPriorityCount = UBound(mOrgPriorityArr) Then
         ReDim Preserve mOrgPriorityArr(mOrgPriorityCount + 50)
      End If
   Loop
End Sub
Public Sub DeleteOrgPriorityByOrgId(OrgId As Long)

   Client.Server.DeleteOrgPriorityByOrgId OrgId
End Sub
Public Sub SaveOrgPriority(OrgId As Long, PriorityId As Integer, PriorityDefault As Boolean)

   Client.Server.SaveOrgPriority OrgId, PriorityId, PriorityDefault
End Sub
Public Sub GetOrgPriorityFromId(OrgPriority As clsOrgPriority, OrgId As Long, PriorityId As Integer)

   Dim I As Integer
   
   Set OrgPriority = Nothing
   For I = 0 To mOrgPriorityCount - 1
      If mOrgPriorityArr(I).PriorityId = PriorityId And mOrgPriorityArr(I).OrgId = OrgId Then
         Set OrgPriority = mOrgPriorityArr(I)
         Exit Sub
      End If
   Next I
End Sub
Public Sub FillCombo(Cbo As ComboBox, ByVal OrgId As Long, SelectedId As Integer, SetDefault As Boolean)

   Dim I As Integer
   Dim NewIdx As Integer
   Dim SelectedIndex As Integer
   Dim DefaultIndex As Integer
   Dim CurrId As Integer
   Dim Org As clsOrg
   
   Cbo.Clear
   If OrgId > 0 Then
      Do
         Client.OrgMgr.GetOrgFromId Org, OrgId
         If Not Org Is Nothing Then
            FillComboForOrg Cbo, Org, SelectedId, SetDefault
            OrgId = Org.OrgParent
            If OrgId <= 0 Then
               Exit Do
            End If
         Else
            Exit Do
         End If
      Loop Until Cbo.ListCount > 0
   End If
   
   If Cbo.ListCount = 0 Then
      SelectedIndex = -1
      For I = 0 To mPriorityCount - 1
         CurrId = mPriorityArr(I).PriorityId
         
         Cbo.AddItem mPriorityArr(I).PriortyText
         NewIdx = Cbo.NewIndex
         Cbo.ItemData(NewIdx) = CurrId
         
         If CurrId = SelectedId Then
            SelectedIndex = NewIdx
         End If
         If CurrId = Client.SysSettings.DictInfoDefaultPriorityId Then
            DefaultIndex = NewIdx
         End If
      Next I
      If Cbo.ListCount > 0 Then                 'Set selected only if list is not empty
         If SelectedIndex >= 0 Then             'Did we detect one item to select?
            Cbo.ListIndex = SelectedIndex
         Else
            If SetDefault Then                  'Should we look for a default
               If DefaultIndex >= 0 Then        'Did we found a default?
                  Cbo.ListIndex = DefaultIndex
               Else
                  Cbo.ListIndex = 0             'Take first entry
               End If
            End If
         End If
      End If
   End If
End Sub
Public Sub FillComboForOrg(Cbo As ComboBox, Org As clsOrg, SelectedId As Integer, SetDefault As Boolean)

   Dim I As Integer
   Dim J As Integer
   Dim NewIdx As Integer
   Dim SelectedIndex As Integer
   Dim DefaultIndex As Integer
   Dim CurrId As Integer
   
   SelectedIndex = -1
   DefaultIndex = -1
   
   For I = 0 To mOrgPriorityCount - 1
      If mOrgPriorityArr(I).OrgId = Org.OrgId Then
         CurrId = mOrgPriorityArr(I).PriorityId
         For J = 0 To mPriorityCount - 1
            If CurrId = mPriorityArr(J).PriorityId Then
               Cbo.AddItem mPriorityArr(J).PriortyText
               NewIdx = Cbo.NewIndex
               Cbo.ItemData(NewIdx) = CurrId
               
               If CurrId = SelectedId Then
                  SelectedIndex = NewIdx
               End If
               If mOrgPriorityArr(I).Def Then
                  DefaultIndex = NewIdx
               'Else
               '   If DefaultIndex < 0 Then
               '      If CurrId = Client.SysSettings.DictInfoDefaultPriorityId Then
               '         DefaultIndex = NewIdx
               '      End If
               '   End If
               End If
            End If
         Next J
      End If
   Next I
   
   If Cbo.ListCount > 0 Then                 'Set selected only if list is not empty
      If SelectedIndex >= 0 Then             'Did we detect one item to select?
         Cbo.ListIndex = SelectedIndex
      Else
         If SetDefault Then                  'Should we look for a default
            If DefaultIndex >= 0 Then        'Did we found a default?
               Cbo.ListIndex = DefaultIndex
            'Else
            '   Cbo.ListIndex = 0             'Take first entry
            End If
         End If
      End If
   End If
End Sub

Public Function IndexFromId(Id As Long) As Integer

   Dim I As Integer
 
   For I = 0 To mPriorityCount - 1
      If mPriorityArr(I).PriorityId = Id Then
         IndexFromId = I
         Exit For
      End If
   Next I
End Function

Public Function IdFromIndex(Index As Integer) As Long

   IdFromIndex = mPriorityArr(Index).PriorityId
End Function
Public Function GetFromIndex(Priority As clsPriority, Idx As Integer)

   Set Priority = mPriorityArr(Idx)
End Function
Public Function GetFromId(Priority As clsPriority, Id As Long)

   Dim Idx As Integer
   
   Idx = IndexFromId(Id)
   Set Priority = mPriorityArr(Idx)
End Function
Public Function TextFromId(Id As Long) As String

   Dim I As Integer
 
   For I = 0 To mPriorityCount - 1
      If mPriorityArr(I).PriorityId = Id Then
         TextFromId = mPriorityArr(I).PriortyText
         Exit For
      End If
   Next I
End Function
Public Function TextFromIndex(Index As Integer) As String

   TextFromIndex = mPriorityArr(Index).PriortyText
End Function

Public Property Get Count() As Integer

   Count = mPriorityCount
End Property
Public Function DefPriorityIdForOrg(OrgId As Long) As Long

   Dim Res As Long
   Dim Org As clsOrg

   Res = -1
   Do While OrgId >= 0
      Client.OrgMgr.GetOrgFromId Org, OrgId
      If Not Org Is Nothing Then
         Res = DefPriorityIdForOrgInt(OrgId)
         If Res >= 0 Then
            Exit Do
         End If
         OrgId = Org.OrgParent
         If OrgId <= 0 Then
            Exit Do
         End If
      Else
         Exit Do
      End If
   Loop
   
   DefPriorityIdForOrg = Res
End Function

Private Function DefPriorityIdForOrgInt(OrgId As Long) As Long

   Dim I As Integer
   Dim Res As Long
   
   Res = -1
   For I = 0 To mOrgPriorityCount - 1
      If mOrgPriorityArr(I).OrgId = OrgId Then
         Res = mOrgPriorityArr(I).PriorityId
         If mOrgPriorityArr(I).Def Then Exit For
      End If
   Next I
   
   DefPriorityIdForOrgInt = Res
End Function
Private Sub Class_Initialize()

   ReDim mPriorityArr(20)
   ReDim mOrgPriorityArr(50)
End Sub
