VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Event OnOpenDictation(DictId As String, _
                             PatId As String, _
                             PatName As String, _
                             ExtDictType As String, _
                             ExtOrg As String, _
                             ExtPriority As String, _
                             AuthorShortName As String, _
                             AuthorLongName As String)
Public Event OnCloseDictation(DictId As String, _
                             PatId As String, _
                             PatName As String, _
                             ExtDictType As String, _
                             ExtOrg As String, _
                             ExtPriority As String, _
                             AuthorShortName As String, _
                             AuthorLongName As String)
Public Event OnNewDictation(DictId As String, _
                             PatId As String, _
                             PatName As String, _
                             ExtDictType As String, _
                             ExtOrg As String, _
                             ExtPriority As String, _
                             AuthorShortName As String, _
                             AuthorLongName As String)
Public Event OnCreateDictation()
Public Event OnLogon()
Public Event OnLogout()
Public Event OnOrgChanged()

Private WithEvents mfrmMain As frmMain
Attribute mfrmMain.VB_VarHelpID = -1

Private Const ModuleName = "ClientAPI"
Private AccessCodeOk  As Boolean
Private mExtSystem As String

Public Function Init(ExtSystem As String, AccessCode As String) As Long

   Const FuncName As String = "Init"
   
   Dim Ret As Long

   On Error GoTo Init_Err
   'If Not InitAtEveryFunction(FuncName, ExtSystem, "", False) Then Exit Function
   
   AccessCodeOk = (AccessCode = "dsfkkd8jd,.,sdf88h3%&%&¤iyt")
   If Not AccessCodeOk Then Exit Function
   
   mExtSystem = ExtSystem
   If StartUpFormMainIsLoaded > 0 Then
      Set mfrmMain = frmMain
   End If
   
Init_Exit:
   'Init = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
Init_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function Start(Server As String, _
                      Database As String, _
                      UserLoginName As String, _
                      Password As String, _
                      ExtPassword As String) As Long

   Const FuncName As String = "Start"
   
   Dim Ret As Long

   On Error GoTo Start_Err
   'If Not InitAtEveryFunction(FuncName, Server & "/" & Database, UserLoginName, False) Then Exit Function
   
   If StartUpFormMainIsLoaded = 0 Then
      StartUpServer = Server
      StartUpDatabase = Database
      StartUpUserLoginName = UserLoginName
      StartUpPassword = Password
      StartUpExtPassword = ExtPassword
      StartUpExtSystem = mExtSystem
      On Error Resume Next
      Load frmMain
      On Error GoTo Start_Err
      If StartUpLoginResult = 0 Then
         Set mfrmMain = frmMain
         frmMain.ShowInForeground
         Ret = 0
      Else
         Ret = StartUpLoginResult + 10
      End If
   Else
      If (Len(Server) = 0 Or UCase$(Server) = UCase$(Client.Server.Server)) And _
         (Len(Database) = 0 Or UCase$(Database) = UCase$(Client.Server.Database)) And _
         UCase$(UserLoginName) = UCase$(Client.User.LoginName) Then
         Ret = 1
      Else
         Ret = 2
      End If
      frmMain.ShowInForeground
   End If
Start_Exit:
   'Start = FinishAtEveryFunction(FuncName, Ret)
   Start = Ret
   Exit Function
   
Start_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function ClientClose() As Integer

   Const FuncName As String = "ClientClose"
   
   Dim Ret As Long

   On Error GoTo ClientClose_Err
   If Not InitAtEveryFunction(FuncName) Then Exit Function
   
   'Body
   Unload frmMain
   
ClientClose_Exit:
   ClientClose = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
ClientClose_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function CurrPat(PatId1 As String, _
                        PatId2 As String, _
                        PatName As String, _
                        ExtDictType As String, _
                        ExtOrg As String, _
                        ExtPriority As String) As Long

   Const FuncName As String = "CurrPat"
   
   Dim Ret As Long
   Dim BlankCurrPat As Boolean

   On Error GoTo CurrPat_Err
   If Not InitAtEveryFunction(FuncName, PatId1, PatId2) Then Exit Function
   
   'Body
   
   Dim CurrPatFilter As New clsFilter
   
   BlankCurrPat = Len(PatId1) = 0 And Len(PatId2) = 0
   
   If Len(PatId1) = 0 And Len(PatId2) > 0 Then
      Client.CurrPatient.PatId = PatId2
   Else
      Client.CurrPatient.PatId = PatId1
      Client.CurrPatient.PatId2 = PatId2
   End If
   Client.CurrPatient.PatName = PatName
   Client.CurrPatient.DictTypeId = Client.ExtSystemMgr.GetIntDictTypeId(mExtSystem, ExtDictType)
   Client.CurrPatient.OrgId = Client.ExtSystemMgr.GetIntOrgId(mExtSystem, ExtOrg)
   Client.CurrPatient.PriorityId = Client.ExtSystemMgr.GetIntPriorityId(mExtSystem, ExtPriority)
   
   CurrPatFilter.StatusStart = Client.SysSettings.DictListCurrStatusStart
   CurrPatFilter.StatusEnd = Client.SysSettings.DictListCurrStatusEnd
   CurrPatFilter.Pat.PatId = StringReplace(StringReplace(PatId1, "-", ""), "/", "")
   CurrPatFilter.Pat.PatId2 = StringReplace(StringReplace(PatId2, "-", ""), "/", "")
   
   frmMain.ShowNewCurrPat
   
   Set frmMain.ucDictList.CurrPatientFilter = CurrPatFilter
   
   If Not BlankCurrPat Then
      frmMain.ucOrgTree.PickOrgId 30005
   Else
      frmMain.ucOrgTree.PickOrgId Client.User.HomeOrgId
   End If
   frmMain.Tabs.Tab = 0
   CurrPat = 0
   
CurrPat_Exit:
   CurrPat = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
CurrPat_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function RecNew(PatId As String, _
                       PatName As String, _
                       ExtId As String, _
                       ExtDictType As String, _
                       ExtOrg As String, _
                       ExtPriority As String) As Long

   Const FuncName As String = "RecNew"
   
   Dim Ret As Long

   On Error GoTo RecNew_Err
   If Not InitAtEveryFunction(FuncName, PatId) Then Exit Function
   
   'Body
   
   Dim S As String
      
   On Error Resume Next

   With Client.NewRecInfo
      .ExtSystem = mExtSystem
      .PatId = PatId
      .PatName = PatName
      .ExtDictId = ExtId
      .DictTypeId = Client.ExtSystemMgr.GetIntDictTypeId(mExtSystem, ExtDictType)
      If Len(Client.DictTypeMgr.TextFromId(.DictTypeId)) = 0 Then
         .DictTypeId = 0
      End If
      .PrioId = Client.ExtSystemMgr.GetIntPriorityId(mExtSystem, ExtPriority)
      If Len(Client.PriorityMgr.TextFromId(.PrioId)) = 0 Then
         .PrioId = 0
      End If
      .OrgId = Client.ExtSystemMgr.GetIntOrgId(mExtSystem, ExtOrg)
      If Len(Client.OrgMgr.TextFromId(.OrgId)) = 0 Then
         .OrgId = 0
      End If
      If .OrgId <> 0 Then
         If Not Client.OrgMgr.CheckUserRole(.OrgId, RTAuthor) Then
            .OrgId = 0
         End If
      End If
   End With
   
   frmMain.IsRecNewFromAPI = True
   
RecNew_Exit:
   RecNew = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
RecNew_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function PlayDict(DictId As String, ExtId As String) As Long


   Const FuncName As String = "PlayDict"
   
   Dim Ret As Long

   On Error GoTo PlayDict_Err
   If Not InitAtEveryFunction(FuncName, DictId, ExtId) Then Exit Function
   
   'Body
   Dim DId As Long
   Dim Dict As clsDict
   Dim Res As Integer

   On Error Resume Next
   DId = CLng(DictId)
   If DId = 0 Then
      DId = Client.Server.GetDictIdFromExtDictId(mExtSystem, ExtId)
   End If
   If DId > 0 Then
      Client.PlayDictIdFromAPI = DId
      frmMain.IsPlayFromAPI = True
   End If

PlayDict_Exit:
   PlayDict = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
PlayDict_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function SetVariable(VariableName As String, _
                            VariableValue As String)

   Const FuncName As String = "SetVariable"
   
   Dim Ret As Long

   On Error GoTo SetVariable_Err
   If Not InitAtEveryFunction(FuncName, VariableName, VariableValue) Then GoTo SetVariable_Exit
   
   'Body
   Client.ExtSystemMgr.AddExtVariable mExtSystem, VariableName, VariableValue
   
SetVariable_Exit:
   SetVariable = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
SetVariable_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function GetVariable(VariableName As String, _
                            VariableValue As String) As Long

   Const FuncName As String = "GetVariable"
   
   Dim Ret As Long

   On Error GoTo GetVariable_Err
   If Not InitAtEveryFunction(FuncName, VariableName, VariableValue) Then GoTo GetVariable_Exit
   
   'Body
   If Client.ExtSystemMgr.GetExtVariable(mExtSystem, VariableName, VariableValue) Then
      Ret = 0
   Else
      Ret = -1
   End If
   
GetVariable_Exit:
   GetVariable = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
GetVariable_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function SignDict(DictId As String, _
                         ExtId As String) As Long

   Const FuncName As String = "SignDict"
   
   Dim Ret As Long

   On Error GoTo SignDict_Err
   If Not InitAtEveryFunction(FuncName, DictId, ExtId) Then GoTo SignDict_Exit
   
   'Body
   Dim DId As Long
   Dim Dict As clsDict
   Dim Res As Integer

   On Error Resume Next
   DId = CLng(DictId)
   If DId = 0 Then
      DId = Client.Server.GetDictIdFromExtDictId(mExtSystem, ExtId)
   End If
   If DId > 0 Then
      If Client.DictMgr.CheckOut(Dict, DId, False) = 0 Then
         If Client.OrgMgr.CheckUserRole(Dict.OrgId, RTAuthor) Then
            Dict.StatusId = Transcribed
         End If
         Res = Client.DictMgr.CheckIn(Dict, False)
      End If
   End If

SignDict_Exit:
   SignDict = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
SignDict_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function

End Function
Public Function ClientWindow(Style As Long) As Long

   Const FuncName As String = "ClientWindow"
   
   Dim Ret As Long

   On Error GoTo ClientWindow_Err
   If Not InitAtEveryFunction(FuncName, CStr(Style)) Then GoTo ClientWindow_Exit
   
   'Body
'   If Style >= 0 And Style <= 2 Then
'      frmMain.WindowState = Style
'   Else
'      Ret = -1
'   End If
   If Not IsNull(mfrmMain) Then
      ShowWindow mfrmMain.hwnd, Style
   End If
   
ClientWindow_Exit:
   ClientWindow = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
ClientWindow_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Function NumberOfDictations(PatId As String, _
                                   PatId2 As String, _
                                   Status As Long, _
                                   ByRef DictNum) As Long
                                   
   Const FuncName As String = "NumberOfDictations"
   
   Dim Ret As Long

   On Error GoTo NumberOfDictations_Err
   If Not InitAtEveryFunction(FuncName, PatId & "/" & PatId2, CStr(Status)) Then GoTo NumberOfDictations_Exit
   
   'Body
   '!!!

NumberOfDictations_Exit:
   NumberOfDictations = FinishAtEveryFunction(FuncName, Ret)
   Exit Function
   
NumberOfDictations_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Function
End Function
Public Property Get UserLoginName() As String

   Const FuncName As String = "UserLoginName"
   
   Dim Ret As Long

   On Error GoTo UserLoginName_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo UserLoginName_Exit
   
   'Body
   UserLoginName = Client.User.LoginName
   
UserLoginName_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
UserLoginName_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property
Public Property Get UserShortName() As String

   Const FuncName As String = "UserShortname"
   
   Dim Ret As Long

   On Error GoTo UserShortname_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo UserShortname_Exit
   
   'Body
   UserShortName = Client.User.ShortName

UserShortname_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
UserShortname_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property
Public Property Get UserLongName() As String

   Const FuncName As String = "UserLongName"
   
   Dim Ret As Long

   On Error GoTo UserLongName_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo UserLongName_Exit
   
   'Body
   UserLongName = Client.User.LongName

UserLongName_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
UserLongName_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property
Public Property Get CurrPatId() As String

   Const FuncName As String = "CurrPatId"
   
   Dim Ret As Long

   On Error GoTo CurrPatId_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo CurrPatId_Exit
   
   'Body
   CurrPatId = Client.CurrPatient.PatId

CurrPatId_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
CurrPatId_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property
Public Property Get CurrOrg() As String

   Const FuncName As String = "CurrOrg"
   
   Dim Ret As Long

   On Error GoTo CurrOrg_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo CurrOrg_Exit
   
   'Body
   CurrOrg = frmMain.CurrentOrg
   
CurrOrg_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
CurrOrg_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property
Public Property Get Server() As String

   Const FuncName As String = "Server"
   
   Dim Ret As Long

   On Error GoTo Server_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo Server_Exit
   
   'Body
   Server = Client.Server.Server

Server_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
Server_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property
Public Property Get Database() As String

   Const FuncName As String = "Database"
   
   Dim Ret As Long

   On Error GoTo Database_Err
   If Not InitAtEveryFunction(FuncName) Then GoTo Database_Exit
   
   'Body
   Database = Client.Server.Database

Database_Exit:
   FinishAtEveryFunction FuncName, Ret
   Exit Property
   
Database_Err:
   Ret = ErrorHandler(FuncName, Err)
   Exit Property
End Property

Private Sub Class_Initialize()

   Debug.Print "ClientAPI Initialize"
End Sub

Private Sub Class_Terminate()

   Set mfrmMain = Nothing
End Sub
Public Property Get GUIBusy() As Boolean

   GUIBusy = Not ReadyForApiCalls
End Property

Private Sub mfrmMain_OnOpenDictation(Dict As clsDict)

   Const FuncName As String = "OnOpenDictation"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnOpenDictation(Dict.DictId, Dict.Pat.PatId, Dict.Pat.PatName, "", "", "", Dict.AuthorShortName, Dict.AuthorLongName)
End Sub

Private Sub mfrmMain_OnCloseDictation(Dict As clsDict)

   Const FuncName As String = "OnCloseDictation"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnCloseDictation(Dict.DictId, Dict.Pat.PatId, Dict.Pat.PatName, "", "", "", Dict.AuthorShortName, Dict.AuthorLongName)
End Sub

Private Sub mfrmMain_OnNewDictation(Dict As clsDict)

   Const FuncName As String = "OnNewDictation"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnNewDictation(Dict.DictId, Dict.Pat.PatId, Dict.Pat.PatName, "", "", "", Dict.AuthorShortName, Dict.AuthorLongName)
End Sub

Private Sub mfrmMain_OnCreateDictation()

   Const FuncName As String = "OnCreateDictation"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnCreateDictation
End Sub

Private Sub mfrmMain_OnLogon()

   Const FuncName As String = "OnLogon"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnLogon
End Sub

Private Sub mfrmMain_OnLogout()

   Const FuncName As String = "OnLogout"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnLogout
End Sub

Private Sub mfrmMain_OnOrgChanged()

   Const FuncName As String = "OnOrgChanged"

   AtEveryRaiseEvent FuncName
   RaiseEvent OnOrgChanged
End Sub
Private Function AtEveryRaiseEvent(FuncName As String, Optional ByVal Data1 As String, Optional ByVal Data2 As String) As Boolean

   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, TraceTitle_Event, Data1, Data2
End Function
Private Function InitAtEveryFunction(FuncName As String, Optional ByVal Data1 As String, Optional ByVal Data2 As String, Optional CheckSystem As Boolean = True) As Boolean

   If Not AccessCodeOk Then Exit Function
   
   If CheckSystem Then
      If Client Is Nothing Then
         Set mfrmMain = frmMain
         Load frmMain
      End If
   End If
   
   'If mfrmMain Is Nothing Then
   '   If Not frmMain Is Nothing Then
   '      Set mfrmMain = frmMain
   '   End If
   'End If
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, TraceTitle_Entry, Data1, Data2
   
   If CheckSystem Then
      If Len(mExtSystem) = 0 Then Exit Function
   End If
   InitAtEveryFunction = True
End Function
Private Function FinishAtEveryFunction(FuncName As String, ReturnValue As Long) As Long

   If Client Is Nothing Then Exit Function
   
   FinishAtEveryFunction = ReturnValue
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, TraceTitle_Exit, CStr(ReturnValue)
End Function
Private Function ErrorHandler(FuncName As String, ErrorCode As Long) As Long

   On Error Resume Next   'Allready error, try to do the best we can
   If Client Is Nothing Then Exit Function
   
   ErrorHandler = ErrorCode
   Client.Trace.AddRow Trace_Level_NonFatalErrors, ModuleName, FuncName, TraceTitle_Err, CStr(ErrorCode), Error$(ErrorCode)
End Function

