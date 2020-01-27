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

Public Enum RoleTypeEnum
   RTNotUsed
   RTAuthor
   RTStatistics
   RTHistory
   RTSysAdmin
   RTList
   RTUnlocking
   RTListen
   RTUserAdmin
   RTTranscribe
   RTAuditing
   RTTextEdit
End Enum

Public GroupId       As Long
Public OrgId         As Long
Public Roles         As String
Public Delayed       As Boolean
Public DelayedHours  As Integer

Public Property Let Author(Value As Boolean)

   SetOneRole RTAuthor, Value
End Property
Public Property Get Author() As Boolean

   Author = GetOneRole(RTAuthor)
End Property
Public Property Let Statistics(Value As Boolean)

   SetOneRole RTStatistics, Value
End Property
Public Property Get Statistics() As Boolean

   Statistics = GetOneRole(RTStatistics)
End Property
Public Property Let History(Value As Boolean)

   SetOneRole RTHistory, Value
End Property
Public Property Get History() As Boolean

   History = GetOneRole(RTHistory)
End Property
Public Property Let SysAdmin(Value As Boolean)

   SetOneRole RTSysAdmin, Value
End Property
Public Property Get SysAdmin() As Boolean

   SysAdmin = GetOneRole(RTSysAdmin)
End Property
Public Property Let List(Value As Boolean)

   SetOneRole RTList, Value
End Property
Public Property Get List() As Boolean

   List = GetOneRole(RTList)
End Property
Public Property Let Unlocking(Value As Boolean)

   SetOneRole RTUnlocking, Value
End Property
Public Property Get Unlocking() As Boolean

   Unlocking = GetOneRole(RTUnlocking)
End Property
Public Property Let Listen(Value As Boolean)

   SetOneRole RTListen, Value
End Property
Public Property Get Listen() As Boolean

   Listen = GetOneRole(RTListen)
End Property
Public Property Let UserAdmin(Value As Boolean)

   SetOneRole RTUserAdmin, Value
End Property
Public Property Get UserAdmin() As Boolean

   UserAdmin = GetOneRole(RTUserAdmin)
End Property
Public Property Let Transcriber(Value As Boolean)

   SetOneRole RTTranscribe, Value
End Property
Public Property Get Transcriber() As Boolean

   Transcriber = GetOneRole(RTTranscribe)
End Property
Public Property Let Auditing(Value As Boolean)

   SetOneRole RTAuditing, Value
End Property
Public Property Get Auditing() As Boolean

   Auditing = GetOneRole(RTAuditing)
End Property
Public Property Let TextEditor(Value As Boolean)

   SetOneRole RTTextEdit, Value
End Property
Public Property Get TextEditor() As Boolean

   TextEditor = GetOneRole(RTTextEdit)
End Property

Private Function MakeStringFromEnum(R As RoleTypeEnum) As String

   Select Case R
      Case RTAuthor: MakeStringFromEnum = "A"
      Case RTStatistics: MakeStringFromEnum = "C"
      Case RTHistory: MakeStringFromEnum = "H"
      Case RTSysAdmin: MakeStringFromEnum = "I"
      Case RTList: MakeStringFromEnum = "L"
      Case RTUnlocking: MakeStringFromEnum = "N"
      Case RTListen: MakeStringFromEnum = "R"
      Case RTUserAdmin: MakeStringFromEnum = "S"
      Case RTTranscribe: MakeStringFromEnum = "T"
      Case RTAuditing: MakeStringFromEnum = "U"
      Case RTTextEdit: MakeStringFromEnum = "X"
   End Select
End Function
Private Sub SetOneRole(ByVal R As RoleTypeEnum, Value As Boolean)

   Dim Pos As Integer
   Dim S As String
   
   S = MakeStringFromEnum(R)
   
   Pos = InStr(Roles, S)
   If Value Then
      If Pos = 0 Then
         Roles = Roles & S
      End If
   Else
      If Pos > 0 Then
         Roles = Left$(Roles, Pos - 1) & mId$(Roles, Pos + 1)
      End If
   End If
End Sub

Public Function GetOneRole(ByVal R As RoleTypeEnum) As Boolean

   Dim S As String
   
   If InStr(Roles, "I") > 0 Then
      GetOneRole = True
   Else
      S = MakeStringFromEnum(R)
      GetOneRole = InStr(Roles, S) > 0
   End If
End Function

Public Property Get LoggData() As String

   LoggData = CStr(OrgId) & "/" & CStr(GroupId) & "/" & Roles
End Property

