VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SQLDateStringFormat As String
Public SQLDateDelimiter As String

Public SQLServer As String
Public SQLDatabase As String

Private Db As ADODB.Connection
Private Const DBUserid As String = "ct_public"
Private Const DBPassword As String = "ssdjgfUhssdF&&ghdsj3kjek34094fkkkkkk"
Private Const ct_app_default_password = "kjsd9sd9873hewls983939899%%3_98kdfjjknvy111"
Private Const ct_app_crypto_key = "dkllsahd84h5uig78djdslkf5%%7udfsn#nkgkk"
'Private Const DBUseridAdmin As String = "CTAdmin"
'Private Const DBPasswordAdmin As String = "ssdjgfUhssdFkdhh555"

'Private Settings As clsSettings

Public DictationStorageSoundPath As String
Public DictationStoragePath As String

Private mConnectionString As String
Public StorageOpened As Boolean               'True is Storage opened (connections may be temp closed)
Public AuthMethod As Integer                  '0 ct_user, 1 SQL user, 2 Trusted connection
Private mCryptoAppRolePassword As String        'Encrypted password for app role ct_app
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
   Rs.LockType = adLockReadOnly
   SQL = BuildSQL(SelectFrom, WhereClause, GroupByClause, OrderClause)
   T = Timer
   Rs.Open SQL
   Debug.Print SQL, Format$(Timer - T, "0.00")
   Set Rs.ActiveConnection = Nothing
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

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   
OpenRecordset_Start:

   On Error GoTo OpenRecordset_err
   Db_PrepareConnection
   
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
   End If
   
   T = Timer
   Rs.Open SQL, Db, adOpenDynamic, adLockOptimistic
   Debug.Print SQL, Format$(Timer - T, "0.00")
   Exit Sub
   
OpenRecordset_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume OpenRecordset_Start
   Else
      Error ErrVal
   End If

End Sub
Public Function DictationStorageOpen(Server As String, _
                                     Database As String, _
                                     UserId As String, _
                                     Password As String) As Boolean

   Dim Rs As New ADODB.Recordset
   Dim CS As String
   Dim Provider As String
   Dim Cred As String
   Dim CurrUserid As String
   Dim CurrPassword As String
   Dim UseTrustedInSettings As Boolean

   SQLDateStringFormat = Client.Settings.GetString("DictationStorage", "SQLDateStringFormat", "mm/dd/yyyy")
   SQLDateDelimiter = Client.Settings.GetString("DictationStorage", "SQLDateDelimiter", "'")

   DictationStoragePath = Client.Settings.GetFolder("DictationStorage", "Folder", "")
   DictationStorageSoundPath = Client.Settings.GetFolder("DictationStorage", "SoundFolder", DictationStoragePath)
   
   Provider = Client.Settings.GetString("DictationStorage", "Provider", "sqloledb")
   If Len(Server) = 0 Then
      SQLServer = Client.Settings.GetString("DictationStorage", "Server", "")
   Else
      SQLServer = Server
   End If
      
   If Len(Database) = 0 Then
      SQLDatabase = Client.Settings.GetString("DictationStorage", "Database", "GrundigNova")
   Else
      SQLDatabase = Database
   End If
   
   UseTrustedInSettings = Client.Settings.GetBool("DictationStorage", "TrustedConnection", False)
   
   CS = "Provider=" & Provider
   CS = CS & ";Server=" & SQLServer
   CS = CS & ";Database=" & SQLDatabase
   
   CS = Client.Settings.GetString("DictationStorage", "ConnectionString", CS)
     
   If UseTrustedInSettings Then
      Cred = "Trusted_Connection=yes"
      mConnectionString = CS & ";" & Cred
      AuthMethod = 2
   Else
      Cred = "User Id=" & DBUserid & ";Password=" & DBPassword
      mConnectionString = CS & ";" & Cred
      AuthMethod = 0
   End If
   Db_ConnectionOpen
   DictationStorageOpen = StorageOpened
End Function
Public Sub DictationStorageClose()

   On Error Resume Next
   StorageOpened = False
   Db.ConnectionCloseIfOpened
   DictationStorageSoundPath = ""
End Sub

Public Sub ExecuteCommand(ByRef Cmd As ADODB.Command)

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   
ExecuteCommand_Start:

   On Error GoTo ExecuteCommand_err
   Db_PrepareConnection
   
   Cmd.ActiveConnection = Db

   T = Timer
   Cmd.Execute
   Debug.Print Cmd.CommandText, Format$(Timer - T, "0.00")
   Exit Sub
   
ExecuteCommand_err:
   ErrVal = Err.Number
   If NumberOfTries < 3 Then
      NumberOfTries = NumberOfTries + 1
      Db_TryToRepaireConnection
      Resume ExecuteCommand_Start
   Else
      Error ErrVal
   End If
End Sub
Public Sub ExecuteCommandWithRsResult(ByRef Rs As ADODB.Recordset, ByRef Cmd As ADODB.Command)

   Dim T As Double
   Dim NumberOfTries As Integer
   Dim ErrVal As Long
   
ExecuteCommandWithRsResult_Start:

   On Error GoTo ExecuteCommandWithRsResult_err
   Db_PrepareConnection
   
   Cmd.ActiveConnection = Db
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   'Rs.CursorLocation = adUseClient
   Rs.CursorType = adOpenForwardOnly
   Rs.LockType = adLockOptimistic
   
   Debug.Print Cmd.ActiveConnection.State
   'If Cmd.ActiveConnection.State <> adStateOpen Then Stop '!!!
   T = Timer
   Set Rs = Cmd.Execute
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
Public Sub Db_ConnectionOpen()

   On Error Resume Next
   Db_ConnectionCloseIfOpened
   
   Set Db = New ADODB.Connection
   Db.ConnectionString = mConnectionString
   Db.CommandTimeout = 180
   Db.Open
   StorageOpened = Err = 0

   If StorageOpened Then
      Db_LoginToAppRole
   End If
End Sub
Private Sub Db_PrepareConnection()

   If Db Is Nothing Then
      Db_ConnectionOpen
   ElseIf Db.State = adStateClosed Then
      Db_ConnectionOpen
   End If
End Sub
Private Sub Db_TryToRepaireConnection()

   Dim I As Integer
   
   I = 3
   Do
      MsgWaitObj 2000
      Db_ConnectionOpen
      I = I - 1
   Loop Until I <= 0 Or StorageOpened
End Sub
Public Sub Db_ConnectionCloseIfOpened()

   On Error Resume Next
   If Not Db Is Nothing Then
      Db.Close
   End If
   Set Db = Nothing
End Sub
Private Sub Db_LoginToAppRole()

   If StorageOpened Then
      If Not Db_GetAppInfo() Then
         DictationStorageClose
      End If
      If Not Db_ActivateAppRole() Then
         DictationStorageClose
      End If
   End If
End Sub

Private Function Db_ActivateAppRole() As Boolean

   Dim App_password As String
   
   If Len(mCryptoAppRolePassword) = 0 Then
      App_password = ct_app_default_password
   Else
      App_password = RC4(mCryptoAppRolePassword, ct_app_crypto_key)
   End If

   On Error GoTo ActivateAppRole_Err
   Db.Execute "EXEC sp_setapprole 'ct_app', '" & App_password & "'"
   Db_ActivateAppRole = True
   Exit Function
   
ActivateAppRole_Err:
   Db_ActivateAppRole = False
   Exit Function
End Function
Private Function Db_GetAppInfo() As Boolean

   Dim Rs As ADODB.Recordset
   Dim Cmd As ADODB.Command
   
   Set Cmd = New ADODB.Command
   Cmd.CommandText = "s_AppInfo"
   Cmd.CommandType = adCmdStoredProc

   On Error GoTo GetAppInfo_Err
   ExecuteCommandWithRsResult Rs, Cmd
   If Not Rs.EOF Then
      mCryptoAppRolePassword = nvl(Rs("AppPw"), "")
   End If
   CloseRecordset Rs
   Db_GetAppInfo = True
   Exit Function
   
GetAppInfo_Err:
   Db_GetAppInfo = False
   CloseRecordset Rs
   Exit Function
End Function
Public Sub CloseRecordset(Rs As ADODB.Recordset)

   On Error Resume Next
   Rs.Close
   Set Rs = Nothing
End Sub

