VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDbWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event UIStatusSet(StatusText As String, Busy As Boolean)
Public Event UIStatusSetSub(SubText As String)
Public Event UIStatusProgress(Total As Long, Left As Long)
Public Event UIStatusClear()

Public SQLDateStringFormat As String
Public SQLDateTimeStringFormat As String
Public SQLDateDelimiter As String

Public SQLServer As String
Public SQLDatabase As String

Private Db As ADODB.Connection
'Private Const DBUserid As String = "ct_public"
'Private Const DBPassword As String = "ssdjgfUhssdF&&ghdsj3kjek34094fkkkkkk"
'Private Const ct_app_default_password = "kjsd9sd9873hewls983939899%%3_98kdfjjknvy111"
'Private Const ct_app_crypto_key = "dkllsahd84h5uig78djdslkf5%%7udfsn#nkgkk"
'Private Const DBUseridAdmin As String = "CTAdmin"
'Private Const DBPasswordAdmin As String = "ssdjgfUhssdFkdhh555"
Private Const WSKey1 As String = "sdfgjJks34/6& bsdj"
Private Const WSDummy As String = "kljd8sjkfk%¤hg"
Private Const WSKey2 As String = "KjGs8/hghfdmlfiuoiiojJj6%%%gghaH"

'Private Settings As clsSettings

Public DictationStorageSoundPath As String
Public DictationStoragePath As String

Public StorageOpened As Boolean               'True is Storage opened (connections may be temp closed)
Public AuthMethod As Integer                  '0 ct_user, 1 SQL user, 2 Trusted connection

Private WSClient As GrundigNovaCaller.GrundigNovaCaller
Private mSite As String
Private mHost As String
Private mSessionId As String
Private mSessionKey As String
Private mIniFilePath As String
Private mCtfpath As String
Private mUrl As String
Private mMaxCallSize As Long


Private Enum WSFunctionEnum
   wsNone = 0
   wsOpenDisconnectedRecordset = 1
   wsExecuteCommandWithRsResult = 2
   wsExecuteCommand = 3
   wsOpenFile = 4
   wsUploadFile = 5
   wsStartSession = 6
   wsGetIniFile = 7
   wsGetCtfFile = 8
   wsValidateUser = 9
   wsGetLogoPic = 10
   wsCloseSession = 11
   wsUploadFileBlock = 12
   wsDownloadFileBlock = 13
   wsUploadFileBlockWithType = 14
End Enum

Public Function User_Validate(ByRef LoginName As String, Password As String, ExtSystem As String, ExtPassword As String) As Boolean

   Dim s As String
   Dim UserName As String
   Dim Pos As Integer
   
   LoginName = UCase(Replace(LoginName, "/", "\"))
   
   Pos = InStr(LoginName, "\")
   If Pos > 0 Then
      UserName = mId(LoginName, Pos + 1)
      mSite = UCase(Left(LoginName, Pos - 1))
   Else
      UserName = LoginName
   End If
   
   LoginName = mSite & "\" & UserName
   s = CallWS(wsValidateUser, LoginName & vbTab & Password & vbTab & ExtSystem & vbTab & ExtPassword)
   mSite = s
   If Len(mSite) > 0 Then
      LoginName = mSite & "\" & UserName
   Else
      LoginName = UserName
   End If
   
   Dim IniFileContent As String
   IniFileContent = CallWS(wsGetIniFile, "")
   If Len(IniFileContent) > 0 Then
      mIniFilePath = CreateTempPath & "Nova.ini"
      WriteStringToFile IniFileContent, mIniFilePath
      Client.Settings.IniFilePath = mIniFilePath
   End If
   
   Dim CtfContent As String
   CtfContent = CallWS(wsGetCtfFile, "")
   If Len(CtfContent) > 0 Then
      mCtfpath = CreateTempPath & "Nova.ctf"
      WriteStringToFile CtfContent, mCtfpath
      Client.Texts.Ctfpath = mCtfpath
   End If
   
   'On Error Resume Next
   Dim LogoPicContent As String
   Dim LogoPicPath As String
   LogoPicContent = CallWS(wsGetLogoPic, "")
   If Len(LogoPicContent) > 0 Then
      LogoPicPath = CreateTempPath & "Nova.gif"
      WriteStringToFile LogoPicContent, LogoPicPath
      frmMain.picLogo.Picture = LoadPicture(LogoPicPath)
      KillFileIgnoreError LogoPicPath
   End If
   On Error GoTo 0

   User_Validate = True
End Function
   
'Public Property Get Site() As String
'
'   Site = mSite
'End Property

Private Function CallWS(WSFunction As WSFunctionEnum, Parameter As String) As String

   Dim P As String
   Dim Res As String
   Dim Crypt As String
   Dim Decrypt As String
   Dim Param As String
   Dim ResLength As Long
   Dim ParamLenght As Long
   Dim KeyForThisCall As String
   Dim ErrVal As Long
   
   On Error GoTo CallWS_Err
   
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS1", CStr(WSFunction), CStr(Len(Parameter)), Left(Parameter, 50)

   KeyForThisCall = WSKey1 & mSessionKey & WSKey2
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS11", CStr(WSFunction), CStr(Len(KeyForThisCall)), Left(KeyForThisCall, 50)
   
   Param = Chr(CInt(WSFunction)) & Parameter
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS12", CStr(WSFunction), CStr(Len(Param)), Left(Param, 50)
   ParamLenght = Len(Param)
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS13", CStr(WSFunction), CStr(Len(Param)), Left(Param, 50)
   Param = CStr(ParamLenght) & vbTab & Param
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS14", CStr(WSFunction), CStr(Len(Param)), Left(Param, 50)
   Crypt = RC4(Param, KeyForThisCall)
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS15", CStr(WSFunction), CStr(Len(Crypt)), Left(Crypt, 50)
   P = mSessionId & vbTab & EncodeBase64(Crypt)
   
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS2", CStr(WSFunction), CStr(Len(P)), Left(P, 50)
   
   Res = CallWSint(P, mUrl)
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS3", CStr(WSFunction), CStr(Len(Res)), Left(Res, 50)
   
   Res = DecodeBase64(Res)
   Decrypt = RC4(Res, KeyForThisCall)
   ResLength = 0
   On Error Resume Next
   ResLength = CLng(ConsumeToNextChar(Decrypt, vbTab))
   'Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS4", CStr(WSFunction), CStr(ResLength), Left(Left(Decrypt, ResLength), 50)
   CallWS = Left(Decrypt, ResLength)
   Exit Function
   
CallWS_Err:
   ErrVal = Err.Number
   Client.Trace.AddRow Trace_Level_Full, "CWS", "CWS5", "Err", CStr(ErrVal), Err.description
   Error ErrVal
   Resume Next
End Function

Public Function CallWSint(ByVal P As String, mUrl As String) As String

   Dim s As String
   Dim Res As String
   Dim Id As String
   
   If mMaxCallSize <= 0 Then
      CallWSint = WSClient.WS(P, mUrl)
   Else
      Do
         s = Left(P, mMaxCallSize)
         P = mId(P, mMaxCallSize + 1)
         If Len(s) > 0 Then
            If Len(P) > 0 Then
               Id = WSClient.WSX(Id, s, 2, mUrl)
            Else
               Res = WSClient.WSX(Id, s, 3, mUrl)
            End If
         End If
      Loop While Len(P) > 0
      
      CallWSint = Res
   End If
End Function
Public Sub CopyFileFromSoundTable(DictId As Long, DestPath As String)

   Const MaxBlockSize = 500000
   Dim s As String
   Dim StartIndex As Long  '0 based
   Dim BlockNo As Integer
   
   Do
      s = CallWS(wsDownloadFileBlock, CStr(DictId) & vbTab & StartIndex & vbTab & MaxBlockSize)
      If Len(s) > 0 Then
         BlockNo = BlockNo + 1
         StartIndex = StartIndex + Len(s)
         WriteStringToFile s, DestPath, BlockNo > 1
      End If
   Loop Until Len(s) < MaxBlockSize
End Sub
Public Sub CopyFileToSoundTable(SourcePath As String, DictId As Long, FileName As String, SoundType As String)

   Const MaxBlockSize = 500000
   Dim s As String
   Dim Res As String
   Dim BlockNo As Long
   Dim BlockSize As Long
   Dim FileRest As Long
   Dim FileLength As Long
   
   FileLength = FileLen(SourcePath)
   FileRest = FileLength
   
   Do
      If FileRest > MaxBlockSize Then
         BlockSize = MaxBlockSize
      Else
         BlockSize = FileRest
      End If
      s = ReadStringFromTempFile(SourcePath, BlockSize, FileLength - FileRest + 1)
      FileRest = FileRest - BlockSize
      BlockNo = BlockNo + 1
      If Len(s) > 0 Then
         Res = CallWS(wsUploadFileBlockWithType, CStr(DictId) & vbTab & BlockNo & vbTab & FileLength & vbTab & Len(s) & vbTab & FileName & vbTab & SoundType & vbTab & s)
      End If
   Loop Until FileRest = 0
End Sub
Sub CheckSQLInjection(ByVal SQLstatement As String)

   'Err.Raise 30000   !!!
End Sub
Public Function BuildSQL(SelectFrom As String, WhereClause As String, GroupByClause As String, OrderClause As String)

   Dim Res As String
   
   CheckSQLInjection WhereClause
   CheckSQLInjection GroupByClause
   CheckSQLInjection OrderClause
   
   Res = SelectFrom
   If Len(WhereClause) > 0 Then
      Res = Res & " WHERE " & WhereClause
   End If
   If Len(GroupByClause) > 0 Then
      Res = Res & " GROUP BY " & GroupByClause
   End If
   If Len(OrderClause) > 0 Then
      Res = Res & " ORDER BY " & OrderClause
   End If
   BuildSQL = Res
End Function

Public Property Get Connection() As ADODB.Connection

   Set Connection = Db
End Property

Public Sub OpenDisconnectedRecordset(ByRef Rs As ADODB.Recordset, ByVal SelectFrom As String, ByVal WhereClause As String, ByVal GroupByClause As String, ByVal OrderClause As String)

   Dim s As String
   Dim Fn As String
   Dim T As Double
   Dim SQL As String
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
  
OpenDisconnectedRecordset_Start:

   On Error GoTo OpenDisconnectedRecordset_err
   Db_PrepareConnection
   
   CloseRecordset Rs
   Set Rs = New ADODB.Recordset
   Set Rs.ActiveConnection = Db
   Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   'RS.LockType = adLockReadOnly
   SQL = BuildSQL(SelectFrom, WhereClause, GroupByClause, OrderClause)
   T = Timer
   
   s = CallWS(wsOpenDisconnectedRecordset, SQL)
   Fn = WriteStringToTempFile(s)
   Rs.Open Fn, , , , adCmdFile
   KillFileIgnoreError Fn
   
   Debug.Print SQL, Format$(Timer - T, "0.00")
   Exit Sub
   
OpenDisconnectedRecordset_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume OpenDisconnectedRecordset_Start
   Else
      Error ErrVal
   End If
End Sub
Public Sub OpenRecordset(ByRef Rs As ADODB.Recordset, ByRef SQL As String)

   OpenDisconnectedRecordset Rs, SQL, "", "", ""
End Sub
Public Function DictationStorageOpen(Server As String, _
                                     Database As String, _
                                     UserId As String, _
                                     Password As String) As Boolean
                                     
   Dim WebProtocol As String
   Dim WebServer As String
   Dim WebApp As String
   Dim WebService As String
   Dim WebUrl As String

   WebProtocol = Client.Settings.GetString("Nova", "Protocol", "http")
   WebServer = Client.Settings.GetString("Nova", "Server", "nova.caretalk.se")
   WebApp = Client.Settings.GetString("Nova", "App", "v160")
   WebService = Client.Settings.GetString("Nova", "Service", "Service.asmx")
   WebUrl = Client.Settings.GetString("Nova", "Url", "")
   
   If Len(WebUrl) = 0 Then
      mUrl = WebProtocol & "://" & WebServer & "/" & WebApp & "/" & WebService
   Else
      mUrl = WebUrl
   End If
                                    
   mSite = Client.Settings.GetString("Nova", "Site", Environ("USERDOMAIN"))
   mHost = Client.Settings.GetString("Nova", "Host", GetStationName())
   mMaxCallSize = Client.Settings.GetLong("Nova", "MaxCallSize", 0)
   
   Set WSClient = New GrundigNovaCaller.GrundigNovaCaller
   
   SQLDateStringFormat = Client.Settings.GetString("DictationStorage", "SQLDateStringFormat", "mm/dd/yyyy")
   SQLDateTimeStringFormat = Client.Settings.GetString("DictationStorage", "SQLDateTimeStringFormat", "mm/dd/yyyy hh:nn:ss")
   SQLDateDelimiter = Client.Settings.GetString("DictationStorage", "SQLDateDelimiter", "'")
      
   Dim Res As String
   Res = CallWS(wsStartSession, mSite & vbTab & mHost & vbTab & App.EXEName & vbTab & ApplicationVersion)
   mSessionId = ConsumeToNextChar(Res, vbTab)
   mSessionKey = ConsumeToNextChar(Res, vbTab)
   mSite = ConsumeToNextChar(Res, vbTab)
   'SQLDatabase = mSite
   
   Dim CtfContent As String
   CtfContent = CallWS(wsGetCtfFile, "")
   If Len(CtfContent) > 0 Then
      mCtfpath = CreateTempPath & "Nova.ctf"
      WriteStringToFile CtfContent, mCtfpath
      Client.Texts.Ctfpath = mCtfpath
   End If
      
   StorageOpened = True
   DictationStorageOpen = True
End Function
Public Sub DictationStorageClose()

   On Error Resume Next
   
   CallWS wsCloseSession, ""
   
   StorageOpened = False
   Db.ConnectionCloseIfOpened
   DictationStorageSoundPath = ""
   
   If Len(mIniFilePath) > 0 Then
      KillFileIgnoreError mIniFilePath
   End If
   If Len(mCtfpath) > 0 Then
      KillFileIgnoreError mCtfpath
   End If
   
End Sub

Public Function ExecuteCommand(ByRef Cmd As ADODB.Command) As Long

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   Dim Res As String
   Dim RecordsAffected As String
   Dim ReturnValue As String
   
ExecuteCommand_Start:

   On Error GoTo ExecuteCommand_err
   Db_PrepareConnection
   
   T = Timer

   Res = CallWS(wsExecuteCommand, BuildStringFromCommand(Cmd))
   
   ReturnValue = ConsumeToNextChar(Res, vbTab)
   RecordsAffected = ConsumeToNextChar(Res, vbTab)
   
   If Len(ReturnValue) > 0 Then
      Cmd(0).Value = ReturnValue
   End If
   
   If Len(RecordsAffected) > 0 Then
      ExecuteCommand = CLng(RecordsAffected)
   End If
   
   Debug.Print Cmd.CommandText, Format$(Timer - T, "0.00")
   Exit Function
   
ExecuteCommand_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume ExecuteCommand_Start
   Else
      Error ErrVal
   End If
End Function
Public Sub ExecuteCommandWithRsResult(ByRef Rs As ADODB.Recordset, ByRef Cmd As ADODB.Command)
      
   Dim Fn As String
   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   Dim s As String
   
ExecuteCommandWithRsResult_Start:

   On Error GoTo ExecuteCommandWithRsResult_err
   Db_PrepareConnection
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   'Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockOptimistic
   
   T = Timer
   s = CallWS(wsExecuteCommandWithRsResult, BuildStringFromCommand(Cmd))
   Fn = WriteStringToTempFile(s)
   Rs.Open Fn, , , , adCmdFile
   KillFileIgnoreError Fn
   
   Debug.Print Cmd.CommandText, Format$(Timer - T, "0.00")
   Exit Sub
   
ExecuteCommandWithRsResult_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume ExecuteCommandWithRsResult_Start
   Else
      Error ErrVal
   End If
End Sub
Private Function BuildStringFromCommand(Cmd As ADODB.Command) As String

   Dim I As Integer
   Dim s As String
   Dim Value As String
   
   If Cmd.CommandType = adCmdStoredProc Then
      s = GetCommandText(Cmd.CommandText) & vbTab
   Else
      s = Cmd.CommandText & vbTab
   End If
   s = s & Cmd.CommandType & vbTab
   For I = 0 To Cmd.Parameters.Count - 1
      Select Case Cmd.Parameters(I).Type
         Case adDBTimeStamp, adDate, adDBDate, adDBTime
            Value = Format(Cmd.Parameters(I).Value, "mm/dd/yyyy hh:nn:ss")
            'Value = CStr(Cmd.Parameters(I).Value)
         Case Else
            Value = nvl(Cmd.Parameters(I).Value, "")
      End Select
      s = s & Cmd.Parameters(I).Name & vbTab & Cmd.Parameters(I).Type & vbTab & Cmd.Parameters(I).Direction & vbTab & Cmd.Parameters(I).Size & vbTab & Len(Value) & vbTab & Value & vbTab
   Next I
   BuildStringFromCommand = s
End Function
Public Sub Db_ConnectionOpen()

End Sub
Private Sub Db_PrepareConnection()

End Sub
Private Sub Db_TryToRepaireConnection()

End Sub
Public Sub Db_ConnectionCloseIfOpened()

End Sub
Private Sub Db_LoginToAppRole()

End Sub

Private Function Db_ActivateAppRole() As Boolean

   Db_ActivateAppRole = True
End Function
Private Function Db_GetAppInfo() As Boolean

End Function
Public Sub CloseRecordset(Rs As ADODB.Recordset)

   On Error Resume Next
   Rs.Close
   Set Rs = Nothing
End Sub

Private Sub Class_Terminate()

   Set WSClient = Nothing
   'Set Settings = Nothing
End Sub
Private Function EncodeBase64(s As String) As String

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    Dim arrData() As Byte
    
    If Len(s) = 0 Then
       EncodeBase64 = ""
       Exit Function
    End If
    arrData = StrConv(s, vbFromUnicode)
    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument
    
    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Private Function DecodeBase64(ByVal strData As String) As String

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    If Len(strData) = 0 Then
       DecodeBase64 = ""
       Exit Function
    End If
    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = StrConv(objNode.nodeTypedValue, vbUnicode)
    
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing
End Function
Private Function GetCommandText(s As String) As String

   Dim Pos As Integer
   
   Pos = InStr(s, "call ")
   If Pos > 0 Then
      s = mId$(s, Pos + 5)
   End If
   Pos = InStr(s, "(")
   If Pos > 0 Then
      s = Left(s, Pos - 1)
   Else
      Pos = InStr(s, "}")
      If Pos > 0 Then
         s = Left(s, Pos - 1)
      End If
   End If
   GetCommandText = Trim(s)
End Function
