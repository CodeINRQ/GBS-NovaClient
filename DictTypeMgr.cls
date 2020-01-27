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

Private mDictTypeArr(100) As clsDictType
Private mDictTypeCount As Integer

Public Sub Init()

   If mDictTypeCount = 0 Then
      Client.Server.CreateDictTypeList
      Do While Client.Server.DictTypeListGetNext(mDictTypeArr(mDictTypeCount))
         mDictTypeCount = mDictTypeCount + 1
      Loop
   End If
End Sub
Public Sub FillCombo(Cbo As ComboBox)

   Dim I As Integer
   
   Cbo.Clear
   For I = 0 To mDictTypeCount - 1
      Cbo.AddItem mDictTypeArr(I).DictTypeText
   Next I
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
