VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private EventCol As Collection

Public Sub OnAppEvent(EventName As String)

   Dim E As clsEvent
   
   For Each E In EventCol
      If E.EventName = UCase$(EventName) Then
         OnOneAppEvent E
      End If
   Next E
End Sub

Public Sub OnDictEvent(EventName As String, Dict As clsDict)

   Dim E As clsEvent
   
   For Each E In EventCol
      If E.EventName = UCase$(EventName) Then
         OnOneDictEvent E, Dict
      End If
   Next E
End Sub
Private Sub OnOneAppEvent(E As clsEvent)

   On Error Resume Next
   
   Dim CbString As String
   If Len(E.Clipboard) > 0 Then
      CbString = CreateAppString(E.ExtSystem, E.ClipboardDelimit, E.Clipboard)
      Clipboard.Clear
      Clipboard.SetText CbString
   End If
   
   Dim FilePath As String
   Dim FileContent As String
   If Len(E.FilePath) > 0 Then
      FilePath = CreateAppString(E.ExtSystem, 0, E.FilePath)
      FileContent = CreateAppString(E.ExtSystem, E.FileDelimit, E.FileContent)
      CreateEventFile FilePath, FileContent
   End If

   Dim CmdCompleteLine As String
   Dim CmdPath As String
   Dim ShellRet As Long
   If Len(E.Cmd) > 0 Then
      CmdPath = """" & CreateAppString(E.ExtSystem, 0, E.Cmd) & """"
      CmdCompleteLine = CmdPath & " " & CreateAppString(E.ExtSystem, E.CmdDelimit, E.CmdParams)
      ShellRet = Shell(CmdCompleteLine, E.CmdWindowStyle)
   End If
End Sub

Private Sub OnOneDictEvent(E As clsEvent, D As clsDict)

   On Error Resume Next
   
   Dim CbString As String
   If Len(E.Clipboard) > 0 Then
      CbString = CreateDictString(E.ExtSystem, E.ClipboardDelimit, E.Clipboard, D)
      Clipboard.Clear
      Clipboard.SetText CbString
   End If
   
   Dim FilePath As String
   Dim FileContent As String
   If Len(E.FilePath) > 0 Then
      FilePath = CreateDictString(E.ExtSystem, 0, E.FilePath, D)
      FileContent = CreateDictString(E.ExtSystem, E.FileDelimit, E.FileContent, D)
      CreateEventFile FilePath, FileContent
   End If

   Dim CmdCompleteLine As String
   Dim CmdPath As String
   Dim ShellRet As Long
   If Len(E.Cmd) > 0 Then
      CmdPath = """" & CreateDictString(E.ExtSystem, 0, E.Cmd, D) & """"
      CmdCompleteLine = CmdPath & " " & CreateDictString(E.ExtSystem, E.CmdDelimit, E.CmdParams, D)
      ShellRet = Shell(CmdCompleteLine, E.CmdWindowStyle)
   End If
End Sub
Private Sub CreateEventFile(Path As String, Content As String)

   Dim F As Integer
   
   On Error Resume Next
   Kill Path
   F = FreeFile
   Open Path For Binary Access Write As #F
   Put #F, , Content
   Close #F
End Sub
Private Function CreateAppString(ExtSystem As String, Delimit As Integer, ByVal C As String) As String

   C = CreateVariableString(ExtSystem, Delimit, C)
   C = CreateUserString(Delimit, C)
   C = CreateFormattedString(C)
   CreateAppString = CreateEnvironString(Delimit, C)
End Function
Private Function CreateDictString(ExtSystem As String, Delimit As Integer, ByVal C As String, D As clsDict) As String

   C = ChangeParam(C, Delimit, "DictId", D.DictId)
   C = ChangeParam(C, Delimit, "ExtDictId", D.ExtDictId)
   C = ChangeParam(C, Delimit, "ExtSystem", D.ExtSystem)
   
   C = ChangeParam(C, Delimit, "PatId", D.Pat.PatId)
   C = ChangeParam(C, Delimit, "PatIdFormatted", D.Pat.PatIdFormatted)
   C = ChangeParam(C, Delimit, "PatName", D.Pat.PatName)
   
   C = ChangeParam(C, Delimit, "StatusId", D.StatusId)
   C = ChangeParam(C, Delimit, "StatusText", D.StatusText)
   
   C = ChangeParam(C, Delimit, "DictTypeId", D.DictTypeId)
   C = ChangeParam(C, Delimit, "DictTypeText", D.DictTypeText)
   C = ChangeParam(C, Delimit, "ExtDictType", Client.ExtSystemMgr.GetExtDictType(ExtSystem, D.DictTypeId))

   C = ChangeParam(C, Delimit, "OrgId", D.OrgId)
   C = ChangeParam(C, Delimit, "OrgText", D.OrgText)
   C = ChangeParam(C, Delimit, "ExtOrg", Client.ExtSystemMgr.GetExtOrg(ExtSystem, D.OrgId))

   C = ChangeParam(C, Delimit, "AuthorShortName", D.AuthorShortName)
   C = ChangeParam(C, Delimit, "AuthorLongName", D.AuthorLongName)
   C = ChangeParam(C, Delimit, "TranscriberShortName", D.TranscriberShortName)
   C = ChangeParam(C, Delimit, "TranscriberLongName", D.TranscriberLongName)
   
   C = ChangeParam(C, Delimit, "PriorityId", D.PriorityId)
   C = ChangeParam(C, Delimit, "PriorityText", D.PriorityText)
   C = ChangeParam(C, Delimit, "ExtPriority", Client.ExtSystemMgr.GetExtPriority(ExtSystem, D.PriorityId))
   C = ChangeParam(C, Delimit, "ExpiryDate", Format$(D.ExpiryDate, "ddddd ttttt"))
   
   C = ChangeParam(C, Delimit, "Txt", D.Txt)

   C = CreateVariableString(ExtSystem, Delimit, C)
   C = CreateUserString(Delimit, C)
   C = CreateFormattedString(C)
   CreateDictString = CreateEnvironString(Delimit, C)
End Function
Private Function CreateFormattedString(ByVal C As String) As String

   C = ChangeParam(C, 0, "tab", vbTab)
   C = ChangeParam(C, 0, "lf", vbLf)
   C = ChangeParam(C, 0, "cr", vbCr)
   C = ChangeParam(C, 0, "crlf", vbCrLf)

   CreateFormattedString = C
End Function
Private Function CreateUserString(Delimit As Integer, ByVal C As String) As String

   C = ChangeParam(C, Delimit, "UserLoginName", Client.User.LoginName)
   C = ChangeParam(C, Delimit, "UserShortName", Client.User.ShortName)
   C = ChangeParam(C, Delimit, "UserLongName", Client.User.LongName)

   CreateUserString = C
End Function
Private Function CreateVariableString(ByVal ExtSystem As String, Delimit As Integer, ByVal C As String) As String

   Dim Ev As clsExtVariable
   
   ExtSystem = UCase(ExtSystem)
   For Each Ev In Client.ExtSystemMgr.ExtVariables
      If ExtSystem = Ev.ExtSystem Then
         C = ChangeParam(C, Delimit, Ev.VariableName, Ev.VariableValue)
      End If
   Next Ev

   CreateVariableString = C
End Function
Private Function CreateEnvironString(Delimit As Integer, ByVal C As String) As String

   Dim Pos1 As Integer
   Dim Pos2 As Integer
   Dim Key As String

   Do
      Pos2 = 0
      Pos1 = InStr(C, "%")
      If Pos1 > 0 Then
         Pos2 = InStr(Pos1 + 1, C, "%")
      End If
      If Pos1 > 0 And Pos2 > Pos1 Then
         Key = Mid$(C, Pos1 + 1, Pos2 - Pos1 - 1)
         C = ChangeParam(C, Delimit, Key, Environ(Key))
      End If
   Loop While Pos2 > Pos1
   CreateEnvironString = C
End Function
Private Function ChangeParam(ByVal S As String, ByVal Delimit As Integer, ByVal Param As String, ByVal Value As String) As String

   Dim Pos As Integer
   
   Value = Replace(Value, Chr(Delimit), "")
   If Len(Value) = 0 Or InStr(Value, " ") > 0 And Delimit > 0 Then
      Value = Chr(Delimit) & Value & Chr(Delimit)
   End If
      
   ChangeParam = Replace(S, "%" & Param & "%", Value, 1, -1, vbTextCompare)
End Function
Public Sub Init()

   Dim I As Integer
   
   For I = 0 To Client.ExtSystemMgr.Count - 1
      GetEventFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnOpen"
      GetEventFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnClose"
      GetEventFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnNew"
      GetEventFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnCreate"
      GetEventFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnLogin"
      GetEventFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnLogout"
   Next I
End Sub
Private Sub GetEventFromSettings(ExtSystem As String, EventName As String)

   Dim E As New clsEvent
   Dim EId As String

   E.ExtSystem = UCase$(ExtSystem)
   E.EventName = UCase$(EventName)
   EId = EventId(E.ExtSystem, E.EventName)
   
   E.Clipboard = Client.Settings.GetString(EId, "Clipboard", "")
   E.ClipboardDelimit = Client.Settings.GetLong(EId, "ClipboardDelimit", Asc(""""))
   
   E.FilePath = Client.Settings.GetString(EId, "FilePath", "")
   E.FileContent = Client.Settings.GetString(EId, "FileContent", "")
   E.FileDelimit = Client.Settings.GetLong(EId, "FileDelimit", Asc(""""))

   E.Cmd = Client.Settings.GetString(EId, "Command", "")
   E.CmdParams = Client.Settings.GetString(EId, "CommandParams", "")
   E.CmdWindowStyle = Client.Settings.GetLong(EId, "CommandWindowStyle", 1)
   E.CmdDelimit = Client.Settings.GetLong(EId, "CommandDelimit", Asc(""""))

   E.Obj = Client.Settings.GetString(EId, "Object", "")

   If Len(E.Clipboard) > 0 Or Len(E.FilePath) > 0 Or Len(E.Cmd) > 0 Or Len(E.Obj) > 0 Then
      EventCol.Add E
   Else
      Set E = Nothing
   End If
End Sub
Private Function EventId(ExtSystem As String, EventName As String) As String

   EventId = ExtSystem & " " & EventName
End Function
Private Sub Class_Initialize()

   Set EventCol = New Collection
End Sub

Private Sub Class_Terminate()

   Set EventCol = Nothing
End Sub
