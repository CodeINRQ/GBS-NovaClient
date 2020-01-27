VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroupMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mGroupArr(1000) As clsGroup
Private mGroupCount As Integer
Public Sub DeleteOneUserGroup(UserId As Long, GroupId As Long)

   Client.Server.DeleteOneUserGroup UserId, GroupId
End Sub
Public Sub SaveUserGroup(UserId As Long, GroupId As Long)

   Dim Eno As Long
   
   On Error GoTo SaveUserGroup_Err
   Client.Server.SaveUserGroup UserId, GroupId
   Exit Sub
   
SaveUserGroup_Err:
   Eno = Err.Number
   ErrorHandle "12701", Eno, 12701, "Rollerna kunde inte sparas", False
   Resume Next
End Sub
Public Sub Init()

   If mGroupCount = 0 Then
      Client.Server.CreateGroupList
      Do While Client.Server.GroupListGetNext(mGroupArr(mGroupCount))
         mGroupCount = mGroupCount + 1
      Loop
   End If
End Sub

Public Function IndexFromId(Id As Long) As Integer

   Dim I As Integer
 
   For I = 0 To mGroupCount - 1
      If mGroupArr(I).GroupId = Id Then
         IndexFromId = I
         Exit For
      End If
   Next I
End Function

Public Function IdFromIndex(Index As Integer) As Long

   IdFromIndex = mGroupArr(Index).GroupId
End Function

Public Function TextFromId(Id As Long) As String

   Dim I As Integer
 
   For I = 0 To mGroupCount - 1
      If mGroupArr(I).GroupId = Id Then
         TextFromId = mGroupArr(I).GroupText
         Exit For
      End If
   Next I
End Function

Public Function TextFromIndex(Index As Integer) As String

   TextFromIndex = mGroupArr(Index).GroupText
End Function
Public Property Get Count() As Integer

   Count = mGroupCount
End Property
Public Sub GetGroupFromId(ByRef Grp As clsGroup, GroupId As Long)

   Dim I As Integer
   
   For I = 0 To mGroupCount - 1
      If mGroupArr(I).GroupId = GroupId Then
         Set Grp = mGroupArr(I)
         Exit For
      End If
   Next
End Sub
Public Sub GetGroupFromIndex(ByRef Grp As clsGroup, Index As Integer)

   If Index < mGroupCount Then
      Set Grp = mGroupArr(Index)
   Else
      Set Grp = Nothing
   End If
End Sub
Public Sub SaveGroup(Grp As clsGroup)

   Dim Eno As Long

   On Error GoTo SaveGroup_Err
   
   Client.Server.SaveGroup Grp
   Client.LoggMgr.Insert 1320108, LoggLevel_SysAdmin, 0, Grp.LoggData
   mGroupCount = 0
   Init
   Exit Sub
   
SaveGroup_Err:
   Eno = Err.Number
   ErrorHandle "12702", Eno, 12702, "Gruppen kunde inte sparas", False
   Resume Next
End Sub
