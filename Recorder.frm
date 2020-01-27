VERSION 5.00
Object = "{F6114B2C-1479-43AD-8E1E-2865A319CFB8}#1.0#0"; "SDK2.dll"
Begin VB.Form frmRecordOrg 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrOpenPort 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin DSSSDK2LibCtl.DssRecorderBase DssRecorderBase 
      Left            =   0
      OleObjectBlob   =   "Recorder.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmRecordOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const ModuleName = "frmRecord"

Public Event MicStat(Stat As Long)
Public Event MicStatHW(Hw As String)
Public Event DssOnModeChanged(ByVal mode As Long)
Public Event DssOnPositionChanged(ByVal Position As Long)
Public Event DssOnError(ByVal errString As String, ByVal ErrNum As Long)
Public Event DssOnRecord(ByVal RecStart As Long, ByVal RecEnd As Long, ByVal FileLength As Long, ByVal PeekPerc As Double)
Private m_InRecordMode As Boolean   'to get rid of extra OnRecordEndEvent (bug found 040616)
Private mComPort As String

Public WithEvents Adapter As ADAPTERSERVERLib.AdapterControl
Attribute Adapter.VB_VarHelpID = -1
Public Overwrite As Boolean
Public RecPuse As Boolean
Private m_bCreated As Long

Public Sub SetMicRecordMode(Start As Boolean, Optional Overwrite As Boolean = True, Optional Pause As Boolean = False)

   Const FuncName As String = "SetMicRecordMode"
   
   Dim binsert As Long
   If Overwrite Then
       binsert = 0
   Else
       binsert = 1
   End If
   
   If Pause Then
      binsert = binsert + 2
   End If

   If Start Then
      Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "BeginRecord"
      Adapter.BeginRecordEx binsert, 0
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "UnmuteMicMuteSum"
      DssRecorderBase.UnmuteMicMuteSum
   Else
      Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "EndRecord"
      Adapter.EndRecord
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "MuteMicUnmuteSum"
      DssRecorderBase.MuteMicUnmuteSum
   End If
End Sub
Private Sub Adapter_MicStat(ByVal MicStat As Long, ByVal description As String)

   'MicrophoneStatPrevious = MicrophoneStat
   'MicrophoneStat = MicStat
   'Select Case MicStat
   '   Case 128  'stop
   '      DssRecorderBase.Stop
   '   Case 131  'play
   '      DssRecorderBase.Play
   'End Select
   'RaiseEvent MicStat(MicStat)
   
   Const FuncName As String = "Adapter_MicStat"

   Debug.Print , , , , , MicStat, description

   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "MicStat", CStr(MicStat)
   RaiseEvent MicStat(MicStat)
End Sub

Private Sub Adapter_MicStatDebug(ByVal MicStat As Long, ByVal description As String)

   Const FuncName As String = "Adapter_MicStatDebug"

   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "MicStatDebug", CStr(MicStat) & "," & description
End Sub

Private Sub Adapter_MicStatEx(ByVal MicStat As Long, ByVal micStatClickType As Long, ByVal description As String)

   Const FuncName As String = "Adapter_MicStatEx"

   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "MicStatEx", CStr(MicStat) & "," & CStr(micStatClickType) & "," & description
End Sub

Private Sub Adapter_MicStatHW(ByVal Hw As String, ByVal lastHW As String)

   Const FuncName As String = "Adapter_MicStatHW"

   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "MicStatHW", Hw & "," & lastHW
   RaiseEvent MicStatHW(Hw)
End Sub

Private Sub Adapter_MicStatString(ByVal str As String, ByVal LastString As String)

   Const FuncName As String = "Adapter_MicStatString"

   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "MicStatString", str & "," & LastString
End Sub

Private Sub DssRecorderBase_OnBeginRecord()

   Const FuncName As String = "DssRecorderBase_OnBeginRecord"
   
   Dim Aret As Long

   ' better rec quality
   Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "UnmuteMicMuteSum"
   DssRecorderBase.UnmuteMicMuteSum
    
   ' turn on mic
   Dim binsert As Long
   If Overwrite Then
       binsert = 0
   Else
       binsert = 1
   End If
   
   Debug.Print "On BeginRecord"
   
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "BeginRecordEx", CStr(binsert) & ",0", CStr(Aret)

   m_InRecordMode = True
   'Adapter.BeginRecord
   'DssRecorderBase.UnmuteMicMuteSum
End Sub

Private Sub DssRecorderBase_OnEndRecord()

   Const FuncName As String = "DssRecorderBase_OnEndRecord"

   Debug.Print "On EndRecord"

   If m_InRecordMode Then
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "MuteMicUnmuteSum"
      DssRecorderBase.MuteMicUnmuteSum
      Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "EndRecord"
      On Error Resume Next
      SetMicRecordMode False
   '   Adapter.EndRecord
   '   DssRecorderBase.MuteMicUnmuteSum
      m_InRecordMode = False
   End If
End Sub

Private Sub DssRecorderBase_OnError(ByVal errString As String, ByVal ErrNum As Long)

   Const FuncName As String = "DssRecorderBase_OnError"
   
   Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "OnError", errString & "," & CStr(ErrNum)
   RaiseEvent DssOnError(errString, ErrNum)
End Sub

Private Sub DssRecorderBase_OnModeChanged(ByVal mode As DSSSDK2LibCtl.PlayerMode)

   Const FuncName As String = "DssRecorderBase_OnModeChanged"
   
   'Debug.Print "OnModeChanged" & CStr(mode)
   
   Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "OnModeChanged", CStr(mode)
   RaiseEvent DssOnModeChanged(CLng(mode))
End Sub

Private Sub DssRecorderBase_OnPositionChanged(ByVal Position As Long)

   Const FuncName As String = "DssRecorderBase_OnPositionChanged"
   
   Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "OnPositionChanged", CStr(Position)
   RaiseEvent DssOnPositionChanged(Position)
End Sub

Private Sub DssRecorderBase_OnRecord(ByVal RecStart As Long, ByVal RecEnd As Long, ByVal FileLength As Long, ByVal PeekPerc As Single)

   Const FuncName As String = "DssRecorderBase_OnRecord"
   
   Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "OnRecord", CStr(RecStart) & "," & CStr(RecEnd) & "," & CStr(FileLength) & "," & CStr(PeekPerc)
   RaiseEvent DssOnRecord(RecStart, RecEnd, FileLength, PeekPerc)
End Sub

Private Sub Form_Load()

   Const FuncName As String = "Form_Load"

   Dim AdapterManager As ADAPTERSERVERLib.AdapterControlManager
   Dim PId As Long
   Dim Aret As Long
   
   Set AdapterManager = New ADAPTERSERVERLib.AdapterControlManager
   Set Adapter = AdapterManager.AdapterControl(m_bCreated)
   'MsgBox "m_bCreated " & CStr(m_bCreated)
   Set AdapterManager = Nothing
   PId = GetCurrentProcessId()
   Aret = Adapter.RegisterApp(PId)
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "RegisterApp", CStr(PId), CStr(Aret)
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Const FuncName As String = "Form_Unload"
   
   Dim IsO As Boolean
   Dim PId As Long
   Dim Aret As Long

   IsO = Adapter.IsOpen
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "IsOpen", "", CStr(IsO)
   If IsO Then
      CloseAdapter
   End If
   
   PId = GetCurrentProcessId()
   Aret = Adapter.UnregisterApp(PId)
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "UnregisterApp", CStr(PId), CStr(Aret)
End Sub

Public Sub OpenAdapter(P As String)

   Const FuncName As String = "OpenAdapter"

   If Adapter.IsOpen Then
      CloseAdapter
   End If
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "Open", P
   'MsgBox "P " & P
   Adapter.Open P
End Sub
Private Sub CloseAdapter()

   Const FuncName As String = "CloseAdapter"
   
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "Close"
   Adapter.Close
End Sub

