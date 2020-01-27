VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictTypeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDictTypeArr() As clsDictType
Private mDictTypeCount As Integer

Private mOrgDictTypeArr() As clsOrgDictType
Private mOrgDictTypeCount As Integer
Public Sub SaveDictType(DictType As clsDictType)

   Dim Eno As Long

   On Error GoTo SaveDictType_Err
   
   Client.Server.SaveDictType DictType
   Client.LoggMgr.Insert 1320120, LoggLevel_SysAdmin, 0, DictType.LoggData
   Init
   Exit Sub
   
SaveDictType_Err:
   Eno = Err.Number
   ErrorHandle "12002", Eno, 12002, "Diktattypen kunde inte sparas", False
   Resume Next
End Sub
Public Sub Init()

   mDictTypeCount = 0
   Client.Server.CreateDictTypeList
   Do While Client.Server.DictTypeListGetNext(mDictTypeArr(mDictTypeCount))
      mDictTypeCount = mDictTypeCount + 1
      If mDictTypeCount = UBound(mDictTypeArr) Then
         ReDim Preserve mDictTypeArr(mDictTypeCount + 50)
      End If
   Loop
   
   mOrgDictTypeCount = 0
   Client.Server.CreateOrgDictTypeList
   Do While Client.Server.OrgDictTypeListGetNext(mOrgDictTypeArr(mOrgDictTypeCount))
      mOrgDictTypeCount = mOrgDictTypeCount + 1
      If mOrgDictTypeCount = UBound(mOrgDictTypeArr) Then
         ReDim Preserve mOrgDictTypeArr(mOrgDictTypeCount + 50)
      End If
   Loop
End Sub
Public Sub DeleteOrgDictTypeByOrgId(OrgId As Long)

   Client.Server.DeleteOrgDictTypeByOrgId OrgId
End Sub
Public Sub SaveOrgDictType(OrgId As Long, DictTypeId As Integer, DictTypeDefault As Boolean)

   Client.Server.SaveOrgDictType OrgId, DictTypeId, DictTypeDefault
End Sub
Public Sub GetOrgDictTypeFromId(OrgDictType As clsOrgDictType, OrgId As Long, DictTypeId As Integer)

   Dim I As Integer
   
   Set OrgDictType = Nothing
   For I = 0 To mOrgDictTypeCount - 1
      If mOrgDictTypeArr(I).DictTypeId = DictTypeId And mOrgDictTypeArr(I).OrgId = OrgId Then
         Set OrgDictType = mOrgDictTypeArr(I)
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
      For I = 0 To mDictTypeCount - 1
         CurrId = mDictTypeArr(I).DictTypeId
         
         Cbo.AddItem mDictTypeArr(I).DictTypeText
         NewIdx = Cbo.NewIndex
         Cbo.ItemData(NewIdx) = CurrId
         
         If CurrId = SelectedId Then
            SelectedIndex = NewIdx
         End If
         If CurrId = Client.SysSettings.DictInfoDefaultDictTypeId Then
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
   
   For I = 0 To mOrgDictTypeCount - 1
      If mOrgDictTypeArr(I).OrgId = Org.OrgId Then
         CurrId = mOrgDictTypeArr(I).DictTypeId
         For J = 0 To mDictTypeCount - 1
            If CurrId = mDictTypeArr(J).DictTypeId Then
               Cbo.AddItem mDictTypeArr(J).DictTypeText
               NewIdx = Cbo.NewIndex
               Cbo.ItemData(NewIdx) = CurrId
               
               If CurrId = SelectedId Then
                  SelectedIndex = NewIdx
               End If
               If mOrgDictTypeArr(I).Def Then
                  DefaultIndex = NewIdx
               'Else
               '   If DefaultIndex < 0 Then
               '      If CurrId = Client.SysSettings.DictInfoDefaultDictTypeId Then
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
 
   For I = 0 To mDictTypeCount - 1
      If mDictTypeArr(I).DictTypeId = Id Then
         IndexFromId = I
         Exit For
      End If
   Next I
End Function

Public Function IdFromIndex(Index As Integer) As Long

   IdFromIndex = mDictTypeArr(Index).DictTypeId
End Function
Public Function GetFromIndex(DictType As clsDictType, Idx As Integer)

   Set DictType = mDictTypeArr(Idx)
End Function

Public Function GetFromId(DictType As clsDictType, Id As Long)

   Dim Idx As Integer
   
   Idx = IndexFromId(Id)
   Set DictType = mDictTypeArr(Idx)
End Function

Public Function TextFromId(Id As Long) As String

   Dim I As Integer
 
   For I = 0 To mDictTypeCount - 1
      If mDictTypeArr(I).DictTypeId = Id Then
         TextFromId = mDictTypeArr(I).DictTypeText
         Exit For
      End If
   Next I
End Function

Public Function TextFromIndex(Index As Integer) As String

   TextFromIndex = mDictTypeArr(Index).DictTypeText
End Function
Public Property Get Count() As Integer

   Count = mDictTypeCount
End Property

Private Sub Class_Initialize()

   ReDim mDictTypeArr(50)
   ReDim mOrgDictTypeArr(50)
End Sub
