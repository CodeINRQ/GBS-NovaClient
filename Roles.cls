VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GroupId       As Long
Public OrgId         As Long
Public Roles         As String
Public Delayed       As Boolean
Public DelayedHours  As Integer

Public Property Get LoggData() As String

   LoggData = CStr(OrgId) & "/" & CStr(GroupId) & "/" & Roles
End Property

Public Property Let Supervise(Value As Boolean)

   SetOneRole "S", Value
End Property
Public Property Get Supervise() As Boolean

   Supervise = GetOneRole("S")
End Property
Public Property Let Author(Value As Boolean)

   SetOneRole "A", Value
End Property
Public Property Get Author() As Boolean

   Author = GetOneRole("A")
End Property
Public Property Let Transcriber(Value As Boolean)

   SetOneRole "T", Value
End Property
Public Property Get Transcriber() As Boolean

   Transcriber = GetOneRole("T")
End Property
Public Property Let Listen(Value As Boolean)

   SetOneRole "R", Value
End Property
Public Property Get Listen() As Boolean

   Listen = GetOneRole("R")
End Property
Public Property Let SysAdmin(Value As Boolean)

   SetOneRole "I", Value
End Property
Public Property Get SysAdmin() As Boolean

   SysAdmin = GetOneRole("I")
End Property
Public Property Let List(Value As Boolean)

   SetOneRole "L", Value
End Property
Public Property Get List() As Boolean

   List = GetOneRole("L")
End Property

Private Sub SetOneRole(ByVal R As String, Value As Boolean)

   Dim Pos As Integer
   R = UCase$(R)
   
   Pos = InStr(Roles, R)
   If Value Then
      If Pos = 0 Then
         Roles = Roles & R
      End If
   Else
      If Pos > 0 Then
         Roles = Left$(Roles, Pos - 1) & Mid$(Roles, Pos + 1)
      End If
   End If
End Sub

Private Function GetOneRole(ByVal R As String) As Boolean

   R = UCase$(R)
   GetOneRole = InStr(Roles, R) > 0
End Function

