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

Private Const ModuleName = "EventMgr"

Private EventCol As Collection
Private ThereAreCtCmdFilesToWatch As Boolean
Private ThereAreWindowsToWatch As Boolean

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
   C = ChangeParam(C, Delimit, "Created", Format$(D.Created, "ddddd ttttt"))
   C = ChangeParam(C, Delimit, "Changed", Format$(D.Changed, "ddddd ttttt"))
   
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
         Key = mId$(C, Pos1 + 1, Pos2 - Pos1 - 1)
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
      GetCtCmdFileFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnCtCmdFile"
      GetWatchWindowFromSettings Client.ExtSystemMgr.ExtSystemName(I), "OnWindow"
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
Private Sub GetWatchWindowFromSettings(ExtSystem As String, EventName As String)

   Dim E As New clsEvent
   Dim EId As String

   E.ExtSystem = UCase$(ExtSystem)
   E.EventName = UCase$(EventName)
   EId = EventId(E.ExtSystem, E.EventName)
   
   E.WindowModuleName = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "ModuleName", "")))

   If Len(E.WindowModuleName) > 0 Then
      E.WindowUseTopLevel = Client.Settings.GetBool(EId, "UseTopLevel", True)
      E.WindowIdentControlId = Client.Settings.GetLong(EId, "IdentControlId", -1)
      E.WindowIdentClassName = UCase$(Client.Settings.GetString(EId, "IdentClassName", ""))
      E.WindowIdentCaption = UCase$(Client.Settings.GetString(EId, "IdentCaption", ""))
      E.WindowIdentParse = UCase$(Client.Settings.GetString(EId, "IdentParse", ""))
      E.WindowIdentWindowRectangle = UCase$(Client.Settings.GetString(EId, "IdentWindowRectangle", ""))
      E.WindowIdentClientRectangle = UCase$(Client.Settings.GetString(EId, "IdentClientRectangle", ""))
      
      E.WindowIdControlId = Client.Settings.GetLong(EId, "IdControlId", -1)
      E.WindowIdParse = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "IdParse", "")))
      
      E.WindowNameControlId = Client.Settings.GetLong(EId, "NameControlId", -1)
      E.WindowNameParse = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "NameParse", "")))
      
      E.WindowKeywordControlId = Client.Settings.GetLong(EId, "KeywordControlId", -1)
      E.WindowKeywordParse = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "KeywordParse", "")))
      
      E.WindowOrgControlId = Client.Settings.GetLong(EId, "OrgControlId", -1)
      E.WindowOrgParse = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "OrgParse", "")))
      
      E.WindowPriControlId = Client.Settings.GetLong(EId, "PriControlId", -1)
      E.WindowPriParse = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "PriParse", "")))
      
      E.WindowDictTypeControlId = Client.Settings.GetLong(EId, "DictTypeControlId", -1)
      E.WindowDictTypeParse = UCase$(CreateAppString(ExtSystem, 0, Client.Settings.GetString(EId, "DictTypeParse", "")))

      EventCol.Add E
      ThereAreWindowsToWatch = True
      frmMain.tmrCheckCtCmdFiles.Enabled = True
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
Public Sub CheckForWindow()

   Dim E As clsEvent
   
   If Not ThereAreWindowsToWatch Then Exit Sub
   
   On Error Resume Next
   For Each E In EventCol
      If Len(E.WindowModuleName) > 0 Then
         HandleWatchWindow E
      End If
   Next E
End Sub
Private Sub HandleCtCmdFile(Path As String)

   'MsgBox "New CtCmdFile"
   
   Dim C As New ClientAPI
   Dim Ret As Long
   Dim I As Integer
   Dim Param(10) As String

   On Error Resume Next

   For I = 1 To 10
      Param(I) = GetString("CtCmd", "Param" & CStr(I), "", Path)
   Next I
   C.Init Param(2), API_ACCESS_CODE
   
   Select Case UCase$(Param(1))
      Case "CURRPAT", "CURRPAT2"
         Ret = C.CurrPat(Param(3), _
                         Param(4), _
                         Param(5), _
                         Param(6), _
                         Param(7), _
                         Param(8))
      Case "RECNEW"
         Ret = C.RecNew(Param(3), _
                        Param(4), _
                        Param(5), _
                        Param(6), _
                        Param(7), _
                        Param(8))
      Case "SIGN"
         Ret = C.SignDict(Param(3), _
                          Param(4))
      Case "PLAY"
         Ret = C.PlayDict(Param(3), _
                          Param(4))
      Case "SETVAR"
         Ret = C.SetVariable(Param(3), _
                             Param(4))
   End Select
   Set C = Nothing
End Sub
Private Sub HandleWatchWindow(E As clsEvent)

   Dim C As ClientAPI
   Dim Pos As Integer
   Dim PId As String
   Dim PName As String
   Dim hWnd As Long
   
   hWnd = winGetForegroundWindow()
   If hWnd <> E.WindowLastCheckedHndl Then
      E.WindowLastCheckedHndl = hWnd
      
      If E.WindowUseTopLevel Then
         hWnd = winGetTopLevelWindow(hWnd)
      End If
      
      If hWnd <> E.WindowHndl Then
         If WatchThisWindow(hWnd, E) Then
            E.WindowHndl = hWnd
         End If
      End If
   End If
   If E.WindowHndl <> 0 Then
      HandleWatchOneWindow E
   End If
End Sub
Private Function WatchThisWindow(hWnd As Long, E As clsEvent) As Boolean

   Dim WindowModuleName As String
   Dim WindowCaption As String
   Dim WindowClassName As String
   Dim ControlCaption As String
   Dim ControlWindowRectangle As String
   Dim ControlClientRectangle As String
   Dim TestCaption As String
   Dim hControl As Long

   If Len(E.WindowModuleName) > 0 Then
      WindowModuleName = UCase$(winGetWindowModuleName(hWnd))
      Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, "WatchThisWindow", "ModuleName", WindowModuleName, E.WindowModuleName
      If Right$(WindowModuleName, Len(E.WindowModuleName)) <> E.WindowModuleName Then Exit Function
   End If
   
   If Len(E.WindowIdentClassName) > 0 Then
      WindowClassName = UCase$(winGetClassName(hWnd))
      Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, "WatchThisWindow", "ClassName", WindowClassName, E.WindowIdentClassName
      If E.WindowIdentClassName <> WindowClassName Then Exit Function
   End If
   
   If E.WindowIdentControlId > 0 Then
      If Not FindControlOnWindow(hWnd, E.WindowIdentControlId, ControlCaption, hControl) Then Exit Function
   Else
      ControlCaption = winGetWindowText(hWnd)
   End If
   
   If Len(E.WindowIdentCaption) > 0 Then
      TestCaption = FindString(ControlCaption, E.WindowIdentParse)
      Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, "WatchThisWindow", "IdentCaption", TestCaption, E.WindowIdentCaption
      If UCase$(TestCaption) <> E.WindowIdentCaption Then Exit Function
   End If
   
   If Len(E.WindowIdentWindowRectangle) > 0 Then
      If hControl = 0 Then Exit Function
      
      ControlWindowRectangle = GetWindowRectAsString(hControl)
      Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, "WatchThisWindow", "IdentWindowRectangle", ControlWindowRectangle, E.WindowIdentWindowRectangle
      If E.WindowIdentWindowRectangle <> ControlWindowRectangle Then Exit Function
   End If
   
   If Len(E.WindowIdentClientRectangle) > 0 Then
      If hControl = 0 Then Exit Function
      
      ControlClientRectangle = GetClientRectAsString(hControl)
      Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, "WatchThisWindow", "IdentClientRectangle", ControlClientRectangle, E.WindowIdentClientRectangle
      If E.WindowIdentClientRectangle <> ControlClientRectangle Then Exit Function
   End If
   
   WatchThisWindow = True
End Function
Private Sub HandleWatchOneWindow(E As clsEvent)

   Dim C As ClientAPI
   Dim PatId As String
   Dim PatName As String
   Dim KeyWord As String
   Dim Org As String
   Dim Pri As String
   Dim DictType As String
   
   PatId = RemoveWhiteSpaceFromString(HandleWatchWindowField(E.WindowHndl, E.WindowIdControlId, E.WindowIdParse))
   If IsValidPatientId(PatId) Then
      PatName = HandleWatchWindowField(E.WindowHndl, E.WindowNameControlId, E.WindowNameParse)
      KeyWord = HandleWatchWindowField(E.WindowHndl, E.WindowKeywordControlId, E.WindowKeywordParse)
      Org = HandleWatchWindowField(E.WindowHndl, E.WindowOrgControlId, E.WindowOrgParse)
      Pri = HandleWatchWindowField(E.WindowHndl, E.WindowPriControlId, E.WindowPriParse)
      DictType = HandleWatchWindowField(E.WindowHndl, E.WindowDictTypeControlId, E.WindowDictTypeParse)
      
      If PatId <> E.WindowLastPatid Or PatName <> E.WindowLastName Or KeyWord <> E.WindowLastKeyword Or _
           Org <> E.WindowLastOrg Or Pri <> E.WindowLastPri Or DictType <> E.WindowLastDictType Then
         E.WindowLastPatid = PatId
         E.WindowLastName = PatName
         E.WindowLastKeyword = KeyWord
         E.WindowLastOrg = Org
         E.WindowLastPri = Pri
         E.WindowLastDictType = DictType
         
         Set C = New ClientAPI
         C.Init E.ExtSystem, API_ACCESS_CODE
         C.CurrPat PatId, "", PatName & "¤" & KeyWord, DictType, Org, Pri
         Set C = Nothing
      End If
   End If
End Sub
Private Function IsValidPatientId(PatId As String) As Boolean

   Dim FirstChar As String

   FirstChar = Left$(PatId, 1)
   IsValidPatientId = Len(PatId) >= 10 And FirstChar >= "0" And FirstChar <= "9"
End Function

Private Function HandleWatchWindowField(hWnd As Long, ControlId As Long, Parse As String) As String

   Dim Caption As String
   Static LastCaption As String
   Dim Tmp As String
   Dim hControl As Long
   Dim Ret As String
   
   If ControlId >= 0 Then
      If ControlId = 0 Then
         Caption = winGetWindowText(hWnd)
      Else
         FindControlOnWindow hWnd, ControlId, Caption, hControl
      End If
      Ret = FindString(Caption, Parse)
      If Caption <> LastCaption Then
         LastCaption = Caption
         Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, "HandleWatchWindowField", "Caption", Caption, Ret
      End If
      HandleWatchWindowField = Ret
   End If
End Function
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
Private Function RemoveWhiteSpaceFromString(S As String) As String

   Dim I As Integer
   Dim Res As String
   Dim C As String

   For I = 1 To Len(S)
      C = mId$(S, I, 1)
      If C > " " Then
         Res = Res & C
      End If
   Next I
   RemoveWhiteSpaceFromString = Res
End Function
