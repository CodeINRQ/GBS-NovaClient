VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PatName       As String
Private mPatId1      As String
Private mPatId2      As String

Public Sub Copy(Pat As clsPat)

   Me.PatName = Pat.PatName
   Me.PatId = Pat.PatId
   Me.PatId2 = Pat.PatId2
End Sub
Public Property Get PatIdFormatted() As String

   PatIdFormatted = Client.Custom.FormatPatId(mPatId1)
End Property
Public Property Get PatId2Formatted() As String

   PatId2Formatted = Client.Custom.FormatPatId(mPatId2)
End Property
Public Property Get PatId() As String

   PatId = mPatId1
End Property
Public Property Get PatId2() As String

   PatId2 = mPatId2
End Property
Public Property Let PatId(PId As String)

   mPatId1 = Client.Custom.FormatPatIdForStoring(PId)
End Property
Public Property Let PatId2(PId As String)

   mPatId2 = Client.Custom.FormatPatIdForStoring(PId)
End Property




