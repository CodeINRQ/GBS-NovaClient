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

   If Len(mPatId1) = 12 Then
      PatIdFormatted = Left$(mPatId1, 8) & "-" & mId$(mPatId1, 9)
   ElseIf Len(mPatId1) = 10 Then
      PatIdFormatted = Left$(mPatId1, 6) & "-" & mId$(mPatId1, 7)
   Else
      PatIdFormatted = mPatId1
   End If
End Property
Public Property Get PatId2Formatted() As String

   If Len(mPatId2) = 12 Then
      PatId2Formatted = Left$(mPatId2, 8) & "-" & mId$(mPatId2, 9)
   ElseIf Len(mPatId2) = 10 Then
      PatId2Formatted = Left$(mPatId2, 6) & "-" & mId$(mPatId2, 7)
   Else
      PatId2Formatted = mPatId2
   End If
End Property
Public Property Get PatId() As String

   PatId = mPatId1
End Property
Public Property Get PatId2() As String

   PatId2 = mPatId2
End Property
Public Property Let PatId(PId As String)

   mPatId1 = FormatPatIdForStoring(PId)
End Property
Public Property Let PatId2(PId As String)

   mPatId2 = FormatPatIdForStoring(PId)
End Property

Private Function FormatPatIdForStoring(ByVal S As String) As String

   S = StringReplace(S, "-", "")
   S = StringReplace(S, "/", "")
   S = StringReplace(S, "\", "")
   S = StringReplace(S, ".", "")
   S = StringReplace(S, ",", "")
   FormatPatIdForStoring = S
End Function


