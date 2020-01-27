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

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private EventCol As Collection
Private ThereAreCtCmdFilesToWatch As Boolean

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
Private Function CreateAppString(ExtSystem As String, Delimit As Integer, ByVal c As String) As String

   c = CreateVariableString(ExtSystem, Delimit, c)
   c = CreateUserString(Delimit, c)
   c = CreateFormattedString(c)
   CreateAppString = CreateEnvironString(Delimit, c)
End Function
Private Function CreateDictString(ExtSystem As String, Delimit As Integer, ByVal c As String, D As clsDict) As String

   c = ChangeParam(c, Delimit, "DictId", D.DictId)
   c = ChangeParam(c, Delimit, "ExtDictId", D.ExtDictId)
   c = ChangeParam(c, Delimit, "ExtSystem", D.ExtSystem)
   
   c = ChangeParam(c, Delimit, "PatId", D.Pat.PatId)
   c = ChangeParam(c, Delimit, "PatIdFormatted", D.Pat.PatIdFormatted)
   c = ChangeParam(c, Delimit, "PatName", D.Pat.PatName)
   
   c = ChangeParam(c, Delimit, "StatusId", D.StatusId)
   c = ChangeParam(c, Delimit, "StatusText", D.StatusText)
   c = ChangeParam(c, Delimit, "Created", Format$(D.Created, "ddddd ttttt"))
   c = ChangeParam(c, Delimit, "Changed", Format$(D.Changed, "ddddd ttttt"))
   
   c = ChangeParam(c, Delimit, "DictTypeId", D.DictTypeId)
   c = ChangeParam(c, Delimit, "DictTypeText", D.DictTypeText)
   c = ChangeParam(c, Delimit, "ExtDictType", Client.ExtSystemMgr.GetExtDictType(ExtSystem, D.DictTypeId))

   c = ChangeParam(c, Delimit, "OrgId", D.OrgId)
   c = ChangeParam(c, Delimit, "OrgText", D.OrgText)
   c = ChangeParam(c, Delimit, "ExtOrg", Client.ExtSystemMgr.GetExtOrg(ExtSystem, D.OrgId))

   c = ChangeParam(c, Delimit, "AuthorShortName", D.AuthorShortName)
   c = ChangeParam(c, Delimit, "AuthorLongName", D.AuthorLongName)
   c = ChangeParam(c, Delimit, "TranscriberShortName", D.TranscriberShortName)
   c = ChangeParam(c, Delimit, "TranscriberLongName", D.TranscriberLongName)
   
   c = ChangeParam(c, Delimit, "PriorityId", D.PriorityId)
   c = ChangeParam(c, Delimit, "PriorityText", D.PriorityText)
   c = ChangeParam(c, Delimit, "ExtPriority", Client.ExtSystemMgr.GetExtPriority(ExtSystem, D.PriorityId))
   c = ChangeParam(c, Delimit, "ExpiryDate", Format$(D.ExpiryDate, "ddddd ttttt"))
   
   c = ChangeParam(c, Delimit, "Txt", D.Txt)

   c = CreateVariableString(ExtSystem, Delimit, c)
   c = CreateUserString(Delimit, c)
   c = CreateFormattedString(c)
   CreateDictString = CreateEnvironString(Delimit, c)
End Function
Private Function CreateFormattedString(ByVal c As String) As String

   c = ChangeParam(c, 0, "tab", vbTab)
   c = ChangeParam(c, 0, "lf", vbLf)
   c = ChangeParam(c, 0, "cr", vbCr)
   c = ChangeParam(c, 0, "crlf", vbCrLf)

   CreateFormattedString = c
End Function
Private Function CreateUserString(Delimit As Integer, ByVal c As String) As String

   c = ChangeParam(c, Delimit, "UserLoginName", Client.User.LoginName)
   c = ChangeParam(c, Delimit, "UserShortName", Client.User.ShortName)
   c = ChangeParam(c, Delimit, "UserLongName", Client.User.LongName)

   CreateUserString = c
End Function
Private Function CreateVariableString(ByVal ExtSystem As String, Delimit As Integer, ByVal c As String) As String

   Dim Ev As clsExtVariable
   
   ExtSystem = UCase(ExtSystem)
   For Each Ev In Client.ExtSystemMgr.ExtVariables
      If ExtSystem = Ev.ExtSystem Then
         c = ChangeParam(c, Delimit, Ev.VariableName, Ev.VariableValue)
      End If
   Next Ev

   CreateVariableString = c
End Function
Private Function CreateEnvironString(Delimit As Integer, ByVal c As String) As String

   Dim Pos1 As Integer
   Dim Pos2 As Integer
   Dim Key As String

   Do
      Pos2 = 0
      Pos1 = InStr(c, "%")
      If Pos1 > 0 Then
         Pos2 = InStr(Pos1 + 1, c, "%")
      End If
      If Pos1 > 0 And Pos2 > Pos1 Then
         Key = mId$(c, Pos1 + 1, Pos2 - Pos1 - 1)
         c = ChangeParam(c, Delimit, Key, Environ(Key))
      End If
   Loop While Pos2 > Pos1
   CreateEnvironString = c
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
      GetCtCmdFileFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnCtCmdFile"
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
Private Sub GetCtCmdFileFromSettings(ExtSystem As String, EventName As String)

   Dim E As New clsEvent
   Dim EId As String

   E.ExtSystem = UCase$(ExtSystem)
   E.EventName = UCase$(EventName)
   EId = EventId(E.ExtSystem, E.EventName)
   
   E.CtCmdFile = CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "CtCmdFile", ""))

   If Len(E.CtCmdFile) > 0 Then
      EventCol.Add E
      ThereAreCtCmdFilesToWatch = True
      frmMain.tmrCheckCtCmdFiles.Enabled = True
      KillFileIgnoreError E.CtCmdFile
   Else
      Set E = Nothing
   End If
End Sub
Public Sub CheckForCtCmdFiles()

   Dim E As clsEvent
   Dim Fd As Date
   
   If Not ThereAreCtCmdFilesToWatch Then Exit Sub
   
   On Error Resume Next
   For Each E In EventCol
      If Len(E.CtCmdFile) > 0 Then
         Fd = FileDateTime(E.CtCmdFile)
         If Fd <> E.CtCmdFileDate Then
            E.CtCmdFileDate = Fd
            HandleCtCmdFile E.CtCmdFile
         End If
      End If
   Next E
End Sub
Private Sub HandleCtCmdFile(Path As String)

   'MsgBox "New CtCmdFile"
   
   Dim c As New ClientAPI
   Dim Ret As Long
   Dim I As Integer
   Dim Param(10) As String

   On Error Resume Next

   For I = 1 To 10
      Param(I) = GetString("CtCmd", "Param" & CStr(I), "", Path)
   Next I
   c.Init Param(2), "dsfkkd8jd,.,sdf88h3%&%&¤iyt"
   
   Select Case UCase$(Param(1))
      Case "CURRPAT", "CURRPAT2"
         Ret = c.CurrPat(Param(3), _
                         Param(4), _
                         Param(5), _
                         Param(6), _
                         Param(7), _
                         Param(8))
      Case "RECNEW"
         Ret = c.RecNew(Param(3), _
                        Param(4), _
                        Param(5), _
                        Param(6), _
                        Param(7), _
                        Param(8))
      Case "SIGN"
         Ret = c.SignDict(Param(3), _
                          Param(4))
      Case "PLAY"
         Ret = c.PlayDict(Param(3), _
                          Param(4))
      Case "SETVAR"
         Ret = c.SetVariable(Param(3), _
                             Param(4))
   End Select
   Set c = Nothing
End Sub

Private Function EventId(ExtSystem As String, EventName As String) As String

   EventId = ExtSystem & " " & EventName
End Function
Private Function GetString(Section As String, Entry As String, Default As String, IniFilePath As String) As String

   Dim I As Integer
   Dim S As String * 1024
   Dim Värde As String

   If Len(IniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, Default, S, Len(S), IniFilePath)
      Värde = Left$(S, I)
   Else
      Värde = Default
   End If
   GetString = Värde
End Function
Private Sub Class_Initialize()

   Set EventCol = New Collection
End Sub

Private Sub Class_Terminate()

   frmMain.tmrCheckCtCmdFiles.Enabled = False
   Set EventCol = Nothing
End Sub
