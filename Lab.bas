Attribute VB_Name = "modLab"
Option Explicit

Public Sub CreateManyLoggs(Counter As Integer)

   Do While Counter > 0
      Client.LoggMgr.Insert 3000000, LoggLevel_DictInfo, 0, CStr(Counter)
      Counter = Counter - 1
   Loop
End Sub

Public Sub CreateManyAudit(Counter As Integer)

   Do While Counter > 0
      Client.DictAuditMgr.Insert Counter + 300000, AuditType_CheckOut, 10
      Counter = Counter - 1
   Loop
End Sub


