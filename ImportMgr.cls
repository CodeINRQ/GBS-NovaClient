VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ImportWithDialog(Dict As clsDict, Orgid As Long, DictTypeId As Long)

   If Orgid <= 0 Then
      Orgid = Client.User.HomeOrgId
   End If
   Dict.Orgid = Orgid
   Dict.AuthorId = Client.User.UserId
   Dict.DictTypeIdNoDefault = DictTypeId
   Dict.PriorityId = -1
   
   Dict.Pat.PatId = Client.SysSettings.ImportDefaultId
   Dict.Pat.PatName = Client.SysSettings.ImportDefaultName
   Dict.Txt = Client.SysSettings.ImportDefaultKeyWord
   Dict.Note = Client.SysSettings.ImportDefaultNote
   
   Set mDictForm = New frmDict
   Load mDictForm
   mDictForm.RestoreSettings DictFormSettings
   mDictForm.EditDictation Dict, False
   mDictForm.CloseText(0) = Client.Texts.Txt(1000501, "Radera diktatet")
   mDictForm.CloseTip(0) = Client.Texts.ToolTip(1000501, "Inspelningen kastas!")
   mDictForm.CloseText(1) = Client.Texts.Txt(1000502, "Fortsätt diktera senare")
   mDictForm.CloseTip(1) = Client.Texts.ToolTip(1000502, "Under inspelning")
   mDictForm.CloseText(2) = Client.Texts.Txt(1000503, "Klart för utskrift")
   mDictForm.CloseTip(2) = Client.Texts.ToolTip(1000503, "")
   ShowWindow Me.hWnd, SW_Hide
   mDictForm.Show vbModal
   ShowWindow Me.hWnd, SW_SHOW
   Select Case mDictCloseChoice
      Case 0
         Client.DictFileMgr.KillLocalTempDictationFile Dict
      Case 1
         LastOrgidForNewDictation = Dict.Orgid
         LastDictTypeIdForNewDictation = Dict.DictTypeId
         Dict.StatusId = 20
         If Client.DictMgr.CheckInNew(Dict) Then
            Client.EventMgr.OnDictEvent "OnNew", Dict
            RaiseEvent OnNewDictation(Dict)
         End If
      Case 2
         LastOrgidForNewDictation = Dict.Orgid
         LastDictTypeIdForNewDictation = Dict.DictTypeId
         Dict.StatusId = 30
         If Client.DictMgr.CheckInNew(Dict) Then
            Client.EventMgr.OnDictEvent "OnNew", Dict
            RaiseEvent OnNewDictation(Dict)
         End If
   End Select
   mDictForm.SaveSettings DictFormSettings
   Unload mDictForm
   Set mDictForm = Nothing

End Sub

Public Sub ImportWithSilence(FilePath As String)

End Sub

