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
Public Delayed       As Boolean
Public DelayedHours  As Integer

Private mRoles       As String

Public Property Let Roles(Value As String)

   mRoles = Value
   SetOneRoleInherit RTList
   'If Me.List Then SetOneRoleAllow RTList
   If Me.Listen Then SetOneRoleAllow RTListen
End Property
Public Property Get Roles() As String

   If Me.List Then SetOneRoleAllow RTList
   If Me.Listen Then SetOneRoleAllow RTListen
   Roles = mRoles
End Property

Public Property Let Author(Value As Boolean)

   SetOneRole RTAuthor, Value
End Property
Public Property Get Author() As Boolean

   Author = GetOneRole(RTAuthor)
End Property
Public Property Let AuthorInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTAuthor
   End If
End Property
Public Property Get AuthorInherit() As Boolean

   AuthorInherit = GetOneRoleInherit(RTAuthor)
End Property
Public Property Let Statistics(Value As Boolean)

   SetOneRole RTStatistics, Value
End Property
Public Property Get Statistics() As Boolean

   Statistics = GetOneRole(RTStatistics)
End Property
Public Property Let StatisticsInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTStatistics
   End If
End Property
Public Property Get StatisticsInherit() As Boolean

   StatisticsInherit = GetOneRoleInherit(RTStatistics)
End Property
Public Property Let History(Value As Boolean)

   SetOneRole RTHistory, Value
End Property
Public Property Get History() As Boolean

   History = GetOneRole(RTHistory)
End Property
Public Property Let HistoryInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTHistory
   End If
End Property
Public Property Get HistoryInherit() As Boolean

   HistoryInherit = GetOneRoleInherit(RTHistory)
End Property
Public Property Let SysAdmin(Value As Boolean)

   SetOneRole RTSysAdmin, Value
End Property
Public Property Get SysAdmin() As Boolean

   SysAdmin = GetOneRole(RTSysAdmin)
End Property
Public Property Let SysAdminInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTSysAdmin
   End If
End Property
Public Property Get SysAdminInherit() As Boolean

   SysAdminInherit = GetOneRoleInherit(RTSysAdmin)
End Property
Public Property Let List(Value As Boolean)

   SetOneRole RTList, Value
End Property
Public Property Get List() As Boolean

   Dim Res As Boolean

   'do lazy calculated or for performance reason

   If Me.Author Then
      Res = True
   ElseIf Me.Transcriber Then
      Res = True
   ElseIf Me.Listen Then
      Res = True
   ElseIf Me.TextEditor Then
      Res = True
   ElseIf Me.Auditing Then
      Res = True
   ElseIf Me.Unlocking Then
      Res = True
   End If

   List = Res
End Property
Public Property Let Unlocking(Value As Boolean)

   SetOneRole RTUnlocking, Value
End Property
Public Property Get Unlocking() As Boolean

   Unlocking = GetOneRole(RTUnlocking)
End Property
Public Property Let UnlockingInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTUnlocking
   End If
End Property
Public Property Get UnlockingInherit() As Boolean

   UnlockingInherit = GetOneRoleInherit(RTUnlocking)
End Property
Public Property Let Listen(Value As Boolean)

   SetOneRole RTListen, Value
End Property
Public Property Get Listen() As Boolean

   Listen = GetOneRole(RTListen) Or Me.Author Or Me.Transcriber
End Property
Public Property Let ListenInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTListen
   End If
End Property
Public Property Get ListenInherit() As Boolean

   ListenInherit = GetOneRoleInherit(RTListen)
End Property
Public Property Let UserAdmin(Value As Boolean)

   SetOneRole RTUserAdmin, Value
End Property
Public Property Get UserAdmin() As Boolean

   UserAdmin = GetOneRole(RTUserAdmin)
End Property
Public Property Let UserAdminInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTUserAdmin
   End If
End Property
Public Property Get UserAdminInherit() As Boolean

   UserAdminInherit = GetOneRoleInherit(RTUserAdmin)
End Property
Public Property Let Transcriber(Value As Boolean)

   SetOneRole RTTranscribe, Value
End Property
Public Property Get Transcriber() As Boolean

   Transcriber = GetOneRole(RTTranscribe)
End Property
Public Property Let TranscriberInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTTranscribe
   End If
End Property
Public Property Get TranscriberInherit() As Boolean

   TranscriberInherit = GetOneRoleInherit(RTTranscribe)
End Property
Public Property Let Auditing(Value As Boolean)

   SetOneRole RTAuditing, Value
End Property
Public Property Get Auditing() As Boolean

   Auditing = GetOneRole(RTAuditing)
End Property
Public Property Let AuditingInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTAuditing
   End If
End Property
Public Property Get AuditingInherit() As Boolean

   AuditingInherit = GetOneRoleInherit(RTAuditing)
End Property
Public Property Let TextEditor(Value As Boolean)

   SetOneRole RTTextEdit, Value
End Property
Public Property Get TextEditor() As Boolean

   TextEditor = GetOneRole(RTTextEdit)
End Property
Public Property Let TextEditorInherit(Value As Boolean)

   If Value Then
      SetOneRoleInherit RTTextEdit
   End If
End Property
Public Property Get TextEditorInherit() As Boolean

   TextEditorInherit = GetOneRoleInherit(RTTextEdit)
End Property

Private Function MakeStringFromEnum(r As RoleTypeEnum) As String

   Select Case r
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
Private Sub SetOneRole(ByVal r As RoleTypeEnum, Value As Boolean)

   If Value Then
      SetOneRoleAllow r
   Else
      SetOneRoleDeny r
   End If
End Sub
Private Sub SetOneRoleAllow(ByVal r As RoleTypeEnum)

   Dim Pos As Integer
   Dim s As String
   
   s = MakeStringFromEnum(r)
   
   Pos = InStr(mRoles, s)
   If Pos = 0 Then
      mRoles = mRoles & s
   End If
   
   Pos = InStr(mRoles, LCase(s))
   If Pos > 0 Then
      mRoles = Left$(mRoles, Pos - 1) & mId$(mRoles, Pos + 1)
   End If
End Sub
Private Sub SetOneRoleDeny(ByVal r As RoleTypeEnum)

   Dim Pos As Integer
   Dim s As String
   
   s = MakeStringFromEnum(r)
   
   Pos = InStr(mRoles, LCase(s))
   If Pos = 0 Then
      mRoles = mRoles & LCase(s)
   End If
   
   Pos = InStr(mRoles, s)
   If Pos > 0 Then
      mRoles = Left$(mRoles, Pos - 1) & mId$(mRoles, Pos + 1)
   End If
End Sub
Private Sub SetOneRoleInherit(ByVal r As RoleTypeEnum)

   Dim Pos As Integer
   Dim s As String
   
   s = MakeStringFromEnum(r)
   
   Pos = InStr(mRoles, LCase(s))
   If Pos > 0 Then
      mRoles = Left$(mRoles, Pos - 1) & mId$(mRoles, Pos + 1)
   End If
   
   Pos = InStr(mRoles, s)
   If Pos > 0 Then
      mRoles = Left$(mRoles, Pos - 1) & mId$(mRoles, Pos + 1)
   End If
End Sub

Public Function GetOneRole(ByVal r As RoleTypeEnum) As Boolean

   Dim s As String
   
   If InStr(mRoles, "I") > 0 Then
      GetOneRole = True
   Else
      If r = RTList Then
         GetOneRole = Me.List
      Else
         s = MakeStringFromEnum(r)
         GetOneRole = InStr(mRoles, s) > 0
      End If
   End If
End Function
Private Function GetOneRoleInherit(ByVal r As RoleTypeEnum) As Boolean

   Dim s As String
   
   s = MakeStringFromEnum(r)
   If InStr(mRoles, s) > 0 Then
      GetOneRoleInherit = False
   Else
      If InStr(mRoles, LCase(s)) > 0 Then
         GetOneRoleInherit = False
      Else
         GetOneRoleInherit = True
      End If
   End If
End Function

Public Property Get LoggData() As String

   LoggData = CStr(OrgId) & "/" & CStr(GroupId) & "/" & mRoles
End Property

Public Function AnyAllowRoles() As Boolean

   Dim Res As Boolean
   
   If Me.List Then
      Res = True
   ElseIf Me.History Then
      Res = True
   ElseIf Me.Statistics Then
      Res = True
   ElseIf Me.UserAdmin Then
      Res = True
   End If
   
   AnyAllowRoles = Res
End Function
'To get compatibilty with clients <1.30
Public Sub CleanRolesBeforeSaveRoot()

   Dim I As Integer
   Dim OnlyUppercaseLetters As String
   Dim OneLetter
   
   For I = 1 To Len(mRoles)
      OneLetter = mId$(mRoles, I, 1)
      If OneLetter <= "Z" Then
         OnlyUppercaseLetters = OnlyUppercaseLetters & OneLetter
      End If
   Next I
   mRoles = OnlyUppercaseLetters
End Sub
