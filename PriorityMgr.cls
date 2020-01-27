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

Private mPriorityArr(100) As New clsPriority
Private mPriorityCount As Integer

Public Sub Init()

   If mPriorityCount = 0 Then
      Client.Server.CreatePriorityList
      Do While Client.Server.PriorityListGetNext(mPriorityArr(mPriorityCount))
         mPriorityCount = mPriorityCount + 1
      Loop
   End If
End Sub

Public Sub FillCombo(Cbo As ComboBox)

   Dim I As Integer
   
   Cbo.Clear
   For I = 0 To mPriorityCount - 1
      Cbo.AddItem mPriorityArr(I).PriortyText
   Next I
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
Public Function TextFromIndex(Index As Integer) As String

   TextFromIndex = mPriorityArr(Index).PriortyText
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
Public Function DaysFromIndex(Index As Integer) As Integer

   DaysFromIndex = mPriorityArr(Index).Days
End Function
Public Property Get Count() As Integer

   Count = mPriorityCount
End Property
