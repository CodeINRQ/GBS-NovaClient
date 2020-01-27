VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoggMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum LoggLevelEnum
   LoggLevel_SysFailure = 10
   LoggLevel_SysAdmin = 20
   LoggLevel_SysInfo = 30

   LoggLevel_UserFailure = 40
   LoggLevel_UserAdmin = 50
   LoggLevel_UserInfo = 60
   
   LoggLevel_DictFailure = 70
   LoggLevel_DictAdmin = 80
   LoggLevel_DictInfo = 90
End Enum

Public Sub Insert(LoggId As Long, LoggLevel As Integer, DictId As Long, LoggData As String)

   If Client.SysSettings.LoggEnable Then
      Dim L As clsLogg
      
      Set L = New clsLogg
      L.LoggId = LoggId
      L.LoggLevel = LoggLevel
      L.DictId = DictId
      L.LoggData = LoggData
      
      L.UserId = Client.User.UserId
      L.StationId = Client.Station.Id
      Client.Server.LoggInsert L
      Set L = Nothing
   End If
End Sub
Public Sub CreateList(StartDate As Date, EndDate As Date, StartLevel As Integer, EndLevel As Integer)

   Client.Server.CreateLoggList StartDate, EndDate, StartLevel, EndLevel
End Sub
Public Function GetNext(Logg As clsLogg) As Boolean

   GetNext = Client.Server.LoggListGetNext(Logg)
End Function
