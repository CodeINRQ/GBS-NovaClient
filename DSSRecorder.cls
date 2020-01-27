VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSSRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ModuleName = "DSSRecorder"
Const SettingSection As String = "DSSRecorder"

Public Event GruEvent(EventType As Gru_Event, Data As Long)

Public Enum Gru_RecMode
   GRU_INSERT = 0       ' 0
   GRU_OVERWRITE        ' 1
End Enum

Public Enum Gru_State
   GRU_UNKNOWN = -1  ' -1
   GRU_STOPPED = 0   ' 0
   GRU_PLAY          ' 1
   GRU_RECPAUSED     ' 2
   GRU_REC           ' 3
   GRU_REWIND        ' 4
   GRU_FORWARD       ' 5
   GRU_PORTNOTOPENED ' 6
   GRU_NOTINITILIZED ' 7
End Enum

Public Enum Gru_Event
   GRU_POSCHANGE = 0    ' 0
   GRU_INPUTCHANGE      ' 1
   GRU_OUTPUTCHANGE     ' 2
   GRU_STATECHANGED     ' 3
   GRU_RECMODECHANGED   ' 4
   GRU_BUTTONPRESS      ' 5
   GRU_INUSE            ' 6
   GRU_HWCHANGED        ' 7
End Enum

Public Enum Gru_Button
   GRU_BUT_UNKNOWN = -1    ' -1
   GRU_BUT_PLAY = 0        ' 0
   GRU_BUT_PLAYPAUSE       ' 1
   GRU_BUT_INDEX           ' 2
   GRU_BUT_DICT            ' 3
   GRU_BUT_REWIND          ' 4
   GRU_BUT_FORWARD         ' 5
   GRU_BUT_REC             ' 6
   GRU_BUT_RECPAUSE        ' 7
   GRU_BUT_INSERT          ' 8
End Enum

Public Enum Gru_Error
   GRU_ERR_NOTIMPLEMENTED = 30000   ' 0
   GRU_ERR_UNKNOWNBUTTON            ' 1
   GRU_ERR_UNKNOWNSTATE             ' 2
   GRU_ERR_FILENOTOPENED            ' 3
   GRU_ERR_READONLY                 ' 4
   GRU_ERR_ILLEGALVALUE             ' 5
   GRU_ERR_NOTRECORDING             ' 6
   GRU_ERR_RECORDING                ' 7
   GRU_ERR_FILEALREADYOPENED        ' 8
   GRU_ERR_PORTFAILURE              ' 9
   GRU_ERR_NOHARDWARESUPPORT        '10
   GRU_ERR_NOOBJECT                 '11
   GRU_ERR_FILENOTEXIST             '12
   GRU_ERR_ALREADYINUSE             '13
End Enum

Public Enum Gru_Harware
   GRU_HW_NONE = 0                  ' 0
   GRU_HW_RECORD                    ' 1
   GRU_HW_TYPIST                    ' 2
End Enum

Private Const GruTraceTitleEntry = "Entry"
Private Const GruTraceTitleExit = "Exit"

Private WithEvents frmRecord As frmRecordOrg
Attribute frmRecord.VB_VarHelpID = -1
Private Trace As clsTrace

Private mState As Gru_State
Private mLastReportedState As Gru_State
Private mFastForwarding As Boolean
Private mRewinding As Boolean
Private mButton As Gru_Button

Private mErrCode As Long
Private mErrDesc As String
Private mErrLoc As String

Private mRecMode As Gru_RecMode
Private mPathForOpenedFile As String
Private mFileOpened As Boolean
Private mReadonly As Boolean
Private mPlaySpeed As Long          '500-30000
Private mWindingSpeed As Integer
Private mPlayBackVolume As Double  '0-1
Private mSoundWinding As Boolean
Private mWindingVolume As Double    '0-1
Private mBackspace As Long
Private mCallBackPositionInterval As Long   'millisec
Private mLastReportedPosition As Long
Private mTimeForReportedPosition As Long
Private mRecording As Boolean
Private mHWtype As Gru_Harware
Private mComPort As String
Private mAutoFindPortAlreadyTried As Boolean
Private mIniFilePath As String
Private mAllreadyInit As Boolean
Private mDirty As Boolean
Private mRecorderInUse As Boolean
Public Sub RegisterAndActivateLicens()

   frmRecord.DssRecorderBase.Register
End Sub
Public Function SetMicRecordMode(Start As Integer) As Long

   Const FuncName As String = "SetMicRecordMode"
   
   Dim Ret As Long
   
   On Error GoTo SetMicRecordMode_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, FuncName, GruTraceTitleEntry, CStr(Start)
   
   If Start <> 0 Then
      If mHWtype <> GRU_HW_RECORD Then
         Ret = GRU_ERR_NOHARDWARESUPPORT
      Else
         If mRecorderInUse Then
            Ret = GRU_ERR_ALREADYINUSE
         Else
            SetRecorderInUse True
            frmRecord.SetMicRecordMode True
            Ret = 0
         End If
      End If
   Else
      frmRecord.SetMicRecordMode False
      SetRecorderInUse False
      Ret = 0
   End If
   
SetMicRecordMode_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, FuncName, GruTraceTitleExit, CStr(Ret)
   SetMicRecordMode = Ret
   Exit Function
   
SetMicRecordMode_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetMicRecordMode_Exit
End Function
Private Sub SetRecorderInUse(InUse As Boolean)

   Dim Data As Long
   
   mRecorderInUse = InUse
   If InUse Then
      Data = 1
   Else
      Data = 0
   End If
   RaiseEvent GruEvent(GRU_INUSE, Data)
End Sub
Private Sub Class_Initialize()

   Dim HomeFolder As String
   
   mState = GRU_NOTINITILIZED
   mLastReportedState = mState
   
   Set frmRecord = New frmRecordOrg
End Sub
Public Function Initialize(IniFilePath As String)

   Const FuncName As String = "Initialize"

   
   If mAllreadyInit Then Exit Function
   mAllreadyInit = True
   
   On Error GoTo Initialize_Err
   
     
   Load frmRecord
   
   mRecMode = GRU_INSERT
   mFileOpened = False
   mReadonly = True
   mPlaySpeed = 1000
   mWindingSpeed = 2000
   mPlayBackVolume = 0.5
   mWindingVolume = mPlayBackVolume
   mSoundWinding = True
   mBackspace = 1000
   mCallBackPositionInterval = 100
   mRecording = False
   
   mState = GRU_PORTNOTOPENED
   mComPort = Client.Settings.GetString(SettingSection, "LastUsedComport", "USB")
   OpenPort mComPort
   mState = GRU_STOPPED
   ReportStateChanged
   
Initialize_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit
   Exit Function
   
Initialize_Err:
   ErrorHandler FuncName, Err
   Resume Initialize_Exit
End Function


Public Function CheckLicens(CheckRecordFeature As Boolean) As Boolean

   Const FuncName As String = "CheckLicens"

   Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Play"
   frmRecord.DssRecorderBase.Play
   If CheckRecordFeature Then
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Record", "0"
      frmRecord.DssRecorderBase.Record 0
   End If
   CheckLicens = True
End Function

Public Property Get Dirty() As Boolean

   Dirty = mDirty
End Property
Private Function OpenPort(ComPort As String) As Long

   Const FuncName As String = "OpenPort"
   
   Dim Ret As Long
   Dim I As Integer

   On Error GoTo OpenPort_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ComPort
   
   TryToOpenPort mComPort
   GetHardWare mHWtype
   
OpenPort_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret), CStr(mHWtype)
   OpenPort = Ret
   Exit Function
   
OpenPort_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume OpenPort_Exit
End Function
Public Property Get DeviceName() As String

   If Client.Hw <> GRU_HW_NONE Then
      If frmRecord.Adapter.IsOpen Then
         DeviceName = frmRecord.Adapter.ReceiveString(PCSTAT_ASK_KENNUNG)
      End If
   End If
End Property
Public Property Get DeviceSerialNo() As String

   If Client.Hw <> GRU_HW_NONE Then
      If frmRecord.Adapter.IsOpen Then
         DeviceSerialNo = frmRecord.Adapter.ReceiveString(PCSTAT_ASK_SERNO)
      End If
   End If
End Property
Public Property Get DeviceFirmwareVersion() As String

   If Client.Hw <> GRU_HW_NONE Then
      If frmRecord.Adapter.IsOpen Then
         DeviceFirmwareVersion = frmRecord.Adapter.ReceiveString(PCSTAT_ASK_VERSION)
      End If
   End If
End Property
Private Sub TryToOpenPort(FirstChoicePort As String)

   Const FuncName As String = "TryToOpenPort"
   
   Dim T As Long

   On Error GoTo TryToOpenPort_Err
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, FirstChoicePort
   
   On Error Resume Next
   T = MyGetTickCount()
   Debug.Print frmRecord.Adapter.AdapterType
   'MsgBox "AdapterType " & frmRecord.Adapter.AdapterType
   frmRecord.OpenAdapter FirstChoicePort
   If mHWtype = GRU_HW_NONE And Len(FirstChoicePort) > 0 Then
      frmRecord.OpenAdapter ""
   End If
   'MsgBox "AdapterType " & frmRecord.Adapter.AdapterType
   Debug.Print frmRecord.Adapter.AdapterType
   
TryToOpenPort_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(mHWtype)
   Exit Sub
   
TryToOpenPort_Err:
   ErrorHandler FuncName, Err
   Resume TryToOpenPort_Exit
End Sub
Public Function GetHardWare(ByRef HWtype As Gru_Harware) As Long

   Const FuncName As String = "GetHardWare"
   
   Dim Ret As Long
   Dim IsRHw As Boolean
   Dim IsC As Boolean

   On Error GoTo GetHardWare_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   HWtype = GRU_HW_NONE
   
   IsRHw = frmRecord.Adapter.IsRecHardware
   Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "IsRecHardware", CStr(IsRHw)
   If IsRHw Then
      HWtype = GRU_HW_RECORD
   Else
      IsC = frmRecord.Adapter.IsConnected
      Client.Trace.AddRow Trace_Level_Adapter, ModuleName, FuncName, "IsConnected", CStr(IsC)
      If IsC Then
         HWtype = GRU_HW_TYPIST
      End If
   End If
   
GetHardWare_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret), CStr(HWtype)
   GetHardWare = Ret
   Exit Function
   
GetHardWare_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetHardWare_Exit
End Function
Public Function GetGRUVersion(ByRef Version As Long) As Long

   Const FuncName As String = "GetGRUVersion"
   
   Dim Ret As Long

   On Error GoTo GetGRUVersion_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   Version = CLng(App.Major) * CLng(1000000) + CLng(App.Minor) * CLng(10000) + CLng(App.Revision)
   
GetGRUVersion_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret), CStr(Version)
   GetGRUVersion = Ret
   Exit Function
   
GetGRUVersion_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetGRUVersion_Exit
End Function
Public Function LoadFile(Path As String, ReadOnly As Integer, NewFile As Integer) As Long

   Const FuncName As String = "LoadFile"

   Dim Ret As Long
   Dim L As Long

   On Error GoTo LoadFile_err

   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, Path, CStr(ReadOnly) & " " & CStr(NewFile)
   'TestIfPortOpened
   
   If mFileOpened Then
      Ret = GRU_ERR_FILEALREADYOPENED
   Else
      mDirty = False
      mPathForOpenedFile = Path
      If NewFile <> 0 Then
         On Error Resume Next
         Kill mPathForOpenedFile
         On Error GoTo LoadFile_err
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "CreateFile", mPathForOpenedFile
         frmRecord.DssRecorderBase.CreateFile mPathForOpenedFile
         mDirty = True
      Else
         On Error Resume Next
         L = FileLen(mPathForOpenedFile)
         On Error GoTo LoadFile_err
         If L <= 0 Then
            Ret = GRU_ERR_FILENOTEXIST
            GoTo LoadFile_Exit
         Else
            Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Open", mPathForOpenedFile & "," & CStr(Not ReadOnly)
            frmRecord.DssRecorderBase.Open mPathForOpenedFile, Not ReadOnly
         End If
      End If
      mFileOpened = True
      mReadonly = ReadOnly <> 0
      mRewinding = False
      mFastForwarding = False
   End If
   
LoadFile_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   LoadFile = Ret
   Exit Function
   
LoadFile_err:
   Ret = ErrorHandler(FuncName, Err)
   Resume LoadFile_Exit
End Function
Public Function SaveToFile(Path As String) As Long

   Const FuncName As String = "SaveToFile"

   Dim Ret As Long
   Dim CurrentPos As Long
   Dim CurrentReadOnly As Boolean
   Dim CurrentPath As String

   On Error GoTo SaveToFile_err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, Path
   
   If Not mFileOpened Then
      Ret = GRU_ERR_FILENOTOPENED
   Else
      PlayPauseIntern False
      GetPos CurrentPos
      CurrentReadOnly = mReadonly
      CurrentPath = mPathForOpenedFile
      CloseFile
      If UCase$(Path) <> UCase$(CurrentPath) Then
         On Error Resume Next
         Kill Path
         On Error GoTo SaveToFile_err
         FileCopy CurrentPath, Path
      End If
      LoadFile CurrentPath, CInt(CurrentReadOnly), CInt(False)
      MoveTo CurrentPos
   End If
      
SaveToFile_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SaveToFile = Ret
   Exit Function
   
SaveToFile_err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SaveToFile_Exit
End Function
Public Function CloseFile() As Long

   Const FuncName As String = "CloseFile"
   
   Dim Ret As Long
   
   On Error GoTo CloseFile_Err

   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If Not mFileOpened Then
      Ret = GRU_ERR_FILENOTOPENED
   Else
      mRewinding = False
      mFastForwarding = False
      mFileOpened = False
      mPathForOpenedFile = ""
      mReadonly = True
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Close"
      frmRecord.DssRecorderBase.Close
      mButton = GRU_BUT_UNKNOWN
   End If
   
CloseFile_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   CloseFile = Ret
   Exit Function
   
CloseFile_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume CloseFile_Exit
End Function
Public Function Play() As Long

   Const FuncName As String = "Play"
   
   Dim Ret As Long
   Dim Pos As Long
   
   On Error GoTo Play_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      mRewinding = False
      mFastForwarding = False
      If mState = GRU_STOPPED Then
         If mBackspace > 0 Then
            Pos = frmRecord.DssRecorderBase.Position
            Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Position", CStr(Pos)
            If Pos < mBackspace Then
               Pos = 0
            Else
               Pos = Pos - mBackspace
            End If
            MoveTo Pos
         End If
      End If
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Volume (let)", CStr(mPlayBackVolume)
      frmRecord.DssRecorderBase.Volume = mPlayBackVolume
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Play", CStr(mPlayBackVolume) & ",2"
      frmRecord.DssRecorderBase.Play mPlaySpeed, 2
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
Play_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   Play = Ret
   Exit Function
   
Play_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume Play_Exit
End Function
Public Function Rec() As Long

   Const FuncName As String = "Rec"

   Dim Ret As Long
   Dim Ow As Long
   
   On Error GoTo Rec_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      If Not mReadonly Then
         If mHWtype = GRU_HW_RECORD Then
            mRewinding = False
            mFastForwarding = False
            mRecording = True
            If mRecMode = GRU_OVERWRITE Then
               Ow = 1
            Else
               Ow = 0
            End If
            frmRecord.Overwrite = Ow <> 0
            Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Record", CStr(Ow)
            frmRecord.DssRecorderBase.Pause
            frmRecord.DssRecorderBase.Record Ow
            mDirty = True
            'If mButton = GRU_BUT_PLAYPAUSE Or mButton = GRU_BUT_RECPAUSE Then   'Removed for SDK1 ver 4
               PlayPauseIntern True
            'End If
         Else
            Ret = GRU_ERR_NOHARDWARESUPPORT
         End If
      Else
         Ret = GRU_ERR_READONLY
      End If
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
Rec_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   Rec = Ret
   Exit Function
   
Rec_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume Rec_Exit
End Function
Public Function PlayStop() As Long

   Const FuncName As String = "PlayStop"

   Dim Ret As Long
   
   On Error GoTo PlayStop_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      mRewinding = False
      mFastForwarding = False
      mRecording = False
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
      frmRecord.DssRecorderBase.Stop
      MoveTo 0
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
PlayStop_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   PlayStop = Ret
   Exit Function
   
PlayStop_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume PlayStop_Exit
End Function
Public Function PlayPause() As Long

   PlayPause = PlayPauseIntern(False)
End Function
Private Function PlayPauseIntern(FromButtonOnDevice As Boolean) As Long

   Const FuncName As String = "PlayPauseIntern"

   Dim Ret As Long
   Dim DoPauseInsteadOfStop As Boolean
   
   On Error GoTo PlayPauseIntern_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(FromButtonOnDevice)
   
   If mFileOpened Then
      mRewinding = False
      mFastForwarding = False
      
      If mRecording Then
         If FromButtonOnDevice Then
            DoPauseInsteadOfStop = True
         End If
      Else
         If mBackspace = 0 Then
            DoPauseInsteadOfStop = True
         End If
      End If
      
      If DoPauseInsteadOfStop Then
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Pause"
         frmRecord.DssRecorderBase.Pause
      Else
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
         frmRecord.DssRecorderBase.Stop
      End If
      
      ReportPosition -1, False
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
PlayPauseIntern_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   PlayPauseIntern = Ret
   Exit Function
   
PlayPauseIntern_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume PlayPauseIntern_Exit
End Function
'Public Function RecPause() As Long
'
'   Const FuncName As String = "RecPause"
'
'   Dim Ret As Long
'
'   On Error GoTo RecPause_Err
'
'   Client.Trace.AddRow Trace_Level_FunctionCalls, Modulename, FuncName, GruTraceTitleEntry
'
'   If mFileOpened Then
'      If mRecording Then
'         mRewinding = False
'         mFastForwarding = False
'         frmRecord.DssRecorderBase.Pause
'         ReportPosition -1, False
'      Else
'         Ret = GRU_ERR_NOTRECORDING
'      End If
'   Else
'      Ret = GRU_ERR_FILENOTOPENED
'   End If
'
'RecPause_Exit:
'   Client.Trace.AddRow Trace_Level_FunctionCalls, Modulename, FuncName, GruTraceTitleExit, CStr(Ret)
'   RecPause = Ret
'   Exit Function
'
'RecPause_Err:
'   Ret = ErrorHandler(FuncName, Err)
'   Resume RecPause_Exit
'End Function
Public Function RecResume() As Long

   Const FuncName As String = "RecResume"

   Dim Ret As Long
   
   On Error GoTo RecResume_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      If mRecording Then
         mRewinding = False
         mFastForwarding = False
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Resume"
         frmRecord.DssRecorderBase.Resume
      Else
         Ret = GRU_ERR_NOTRECORDING
      End If
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
RecResume_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   RecResume = Ret
   Exit Function
   
RecResume_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume RecResume_Exit
End Function
Public Function Rewind() As Long

   Const FuncName As String = "Rewind"

   Dim Ret As Long
   Dim Pos As Long
   
   On Error GoTo Rewind_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      mRewinding = True
      mFastForwarding = False
      If mRecording Then
         mRecording = False
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
         frmRecord.DssRecorderBase.Stop
      End If
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
      frmRecord.DssRecorderBase.Stop                       'to force a DssOnModeChanged event every time
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Volume (let)", "0"
      frmRecord.DssRecorderBase.Volume = 0
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Play", CStr(-mWindingSpeed) & ",2"
      frmRecord.DssRecorderBase.Play -mWindingSpeed, 2
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
Rewind_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   Rewind = Ret
   Exit Function
   
Rewind_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume Rewind_Exit
End Function
Public Function FastForward() As Long

   Const FuncName As String = "FastForward"

   Dim Ret As Long
   
   On Error GoTo FastForward_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      mRewinding = False
      mFastForwarding = True
      If mRecording Then
         mRecording = False
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
         frmRecord.DssRecorderBase.Stop
      End If
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
      frmRecord.DssRecorderBase.Stop                       ''to force a DssOnModeChanged event every time
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Volume (let)", CStr(mWindingVolume)
      frmRecord.DssRecorderBase.Volume = mWindingVolume
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Play", CStr(mWindingSpeed) & ",2"
      frmRecord.DssRecorderBase.Play mWindingSpeed, 2
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
FastForward_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   FastForward = Ret
   Exit Function
   
FastForward_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume FastForward_Exit
End Function
Public Function Delete(ByVal FromPos As Long, ByVal ToPos As Long) As Long

   Const FuncName As String = "Delete"

   Dim Ret As Long
   
   On Error GoTo Delete_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(FromPos), CStr(ToPos)
   
   If mFileOpened Then
      If Not mReadonly Then
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "DeleteFromTo", CStr(FromPos) & "," & CStr(ToPos)
         frmRecord.DssRecorderBase.DeleteFromTo FromPos, ToPos
         mDirty = True
         ReportPosition -1, False
      Else
         Ret = GRU_ERR_READONLY
      End If
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
Delete_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   Delete = Ret
   Exit Function
   
Delete_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume Delete_Exit
End Function
Public Function MoveTo(ByVal Pos As Long) As Long

   Const FuncName As String = "MoveTo"
   
   Dim Ret As Long
   
   On Error GoTo MoveTo_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(Pos)
   
   If mFileOpened Then
      mRewinding = False
      mFastForwarding = False
      If mRecording Then
         mRecording = False
         Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Stop"
         frmRecord.DssRecorderBase.Stop
      End If
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Position (let)", CStr(Pos)
      frmRecord.DssRecorderBase.Position = Pos
      ReportPosition Pos, False
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
MoveTo_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   MoveTo = Ret
   Exit Function
   
MoveTo_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume MoveTo_Exit
End Function
Public Function GetPos(ByRef Pos As Long) As Long

   Const FuncName As String = "GetPos"

   Dim Ret As Long
   
   On Error GoTo GetPos_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   If mFileOpened Then
      Pos = frmRecord.DssRecorderBase.Position
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Position (get)", CStr(Pos)
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
GetPos_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Pos)
   GetPos = Ret
   Exit Function
   
GetPos_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetPos_Exit
End Function
Public Function GetLength(ByRef Length As Long) As Long

   Const FuncName As String = "GetLength"
   
   Dim Ret As Long
   
   On Error GoTo GetLength_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   If mFileOpened Then
      Length = frmRecord.DssRecorderBase.Length
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Length", CStr(Length)
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
GetLength_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Length)
   GetLength = Ret
   Exit Function
   
GetLength_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetLength_Exit
End Function
Public Function SetWindingSpeed(ByVal Speed As Integer) As Long

   Const FuncName As String = "SetWindingSpeed"
   
   Dim Ret As Long
   
   On Error GoTo SetWindingSpeed_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(Speed)
   
   If Speed >= 1000 And Speed <= 20000 Then
      mWindingSpeed = Speed
   Else
      Ret = GRU_ERR_ILLEGALVALUE
   End If
   
   
SetWindingSpeed_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetWindingSpeed = Ret
   Exit Function
   
SetWindingSpeed_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetWindingSpeed_Exit
End Function
Public Function SetPlaySpeed(ByVal Speed As Integer) As Long

   Const FuncName As String = "SetPlaySpeed"
   
   Dim Ret As Long
   
   On Error GoTo SetPlaySpeed_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(Speed)
   
   If Speed >= 500 And Speed <= 1500 Then
      mPlaySpeed = Speed
   Else
      Ret = GRU_ERR_ILLEGALVALUE
   End If
   If mState = GRU_PLAY Then
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Play", CStr(mPlaySpeed) & ",2"
      frmRecord.DssRecorderBase.Play mPlaySpeed, 2
   End If
   
   
SetPlaySpeed_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetPlaySpeed = Ret
   Exit Function
   
SetPlaySpeed_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetPlaySpeed_Exit
End Function
Public Function GetWindingSpeed(ByRef Speed As Integer) As Long

   Const FuncName As String = "GetWindingSpeed"
   
   Dim Ret As Long
   
   On Error GoTo GetWindingSpeed_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   Speed = mWindingSpeed
   
GetWindingSpeed_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Speed)
   GetWindingSpeed = Ret
   Exit Function
   
GetWindingSpeed_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetWindingSpeed_Exit
End Function
Public Function GetPlaySpeed(ByRef Speed As Integer) As Long

   Const FuncName As String = "GetPlaySpeed"
   
   Dim Ret As Long
   
   On Error GoTo GetPlaySpeed_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   Speed = mPlaySpeed
   
GetPlaySpeed_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Speed)
   GetPlaySpeed = Ret
   Exit Function
   
GetPlaySpeed_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetPlaySpeed_Exit
End Function
Public Function SetPlayBackVolume(ByVal Volume As Integer) As Long

   Const FuncName As String = "SetPlayBackVolume"
   
   Dim Ret As Long
   
   On Error GoTo SetPlayBackVolume_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(Volume)
   
   If Volume < 0 Or Volume > 100 Then
      Ret = GRU_ERR_ILLEGALVALUE
   Else
      mPlayBackVolume = CDbl(Volume) / 100
      If mSoundWinding Then
         mWindingVolume = mPlayBackVolume
      End If
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Volume (let)", CStr(mPlayBackVolume)
      frmRecord.DssRecorderBase.Volume = mPlayBackVolume
   End If
   
SetPlayBackVolume_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetPlayBackVolume = Ret
   Exit Function
   
SetPlayBackVolume_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetPlayBackVolume_Exit
End Function
Public Function GetPlayBackVolume(ByRef Volume As Integer) As Long

   Const FuncName As String = "GetPlayBackVolume"
   
   Dim Ret As Long
   
   On Error GoTo GetPlayBackVolume_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   Volume = CInt(mPlayBackVolume * 100)
   
GetPlayBackVolume_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Volume)
   GetPlayBackVolume = Ret
   Exit Function
   
GetPlayBackVolume_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetPlayBackVolume_Exit
End Function
Public Function SetRecordMode(ByVal RecMode As Gru_RecMode) As Long

   Const FuncName As String = "SetRecordMode"
   
   Dim Ret As Long
   
   On Error GoTo SetRecordMode_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(RecMode)
   
   If RecMode = GRU_INSERT Or RecMode = GRU_OVERWRITE Then
      mRecMode = RecMode
   Else
      Ret = GRU_ERR_ILLEGALVALUE
   End If
   
SetRecordMode_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetRecordMode = Ret
   Exit Function
   
SetRecordMode_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetRecordMode_Exit
End Function
Public Function GetRecordMode(ByRef RecMode As Gru_RecMode) As Long

   Const FuncName As String = "GetRecordMode"
   
   Dim Ret As Long
   
   On Error GoTo GetRecordMode_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(mRecMode)
   
   RecMode = mRecMode
   
GetRecordMode_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GetRecordMode = Ret
   Exit Function
   
GetRecordMode_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetRecordMode_Exit
End Function
Public Function GetStatus(ByRef State As Gru_State) As Long

   Const FuncName As String = "GetStatus"
   
   Dim Ret As Long
   
   On Error GoTo GetStatus_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(mState)
   
   State = mState
   
GetStatus_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GetStatus = Ret
   Exit Function
   
GetStatus_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetStatus_Exit
End Function
Public Function SetSoundWinding(ByVal SoundWinding As Integer) As Long

   Const FuncName As String = "SetSoundWinding"
   
   Dim Ret As Long
   
   On Error GoTo SetSoundWinding_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(SoundWinding)
   
   mSoundWinding = SoundWinding <> 0
   If mSoundWinding Then
      mWindingVolume = mPlayBackVolume
   Else
      mWindingVolume = 0
   End If
   If mRewinding Or mFastForwarding Then
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Volume (let)", CStr(mWindingVolume)
      frmRecord.DssRecorderBase.Volume = mWindingVolume
   End If
      
SetSoundWinding_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetSoundWinding = Ret
   Exit Function
   
SetSoundWinding_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetSoundWinding_Exit
End Function
Public Function GetSoundWinding(ByRef SoundWinding As Integer) As Long

   Const FuncName As String = "GetSoundWinding"
   
   Dim Ret As Long
   
   On Error GoTo GetSoundWinding_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   SoundWinding = CInt(mSoundWinding)
   
GetSoundWinding_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GetSoundWinding = Ret
   Exit Function
   
GetSoundWinding_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetSoundWinding_Exit
End Function
Public Function SetBackspace(ByVal Backspace As Long) As Long

   Const FuncName As String = "SetBackspace"
   
   Dim Ret As Long
   
   On Error GoTo SetBackspace_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(Backspace)
   
   If Backspace >= 0 And Backspace <= 3000 Then
      mBackspace = Backspace
   Else
      Ret = GRU_ERR_ILLEGALVALUE
   End If
   
SetBackspace_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetBackspace = Ret
   Exit Function
   
SetBackspace_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetBackspace_Exit
End Function
Public Function GetBackspace(ByRef Backspace As Long) As Long

   Const FuncName As String = "GetBackspace"
   
   Dim Ret As Long
   
   On Error GoTo GetBackspace_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   Backspace = mBackspace
   
GetBackspace_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GetBackspace = Ret
   Exit Function
   
GetBackspace_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetBackspace_Exit
End Function
Public Function SetCallBackPositionInterval(ByVal Interval As Long) As Long

   Const FuncName As String = "SetCallBackPositionInterval"
   
   Dim Ret As Long
   
   On Error GoTo SetCallBackPositionInterval_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, CStr(Interval)
   
   If Interval = 0 Or (Interval >= 10 And Interval <= 1000) Then
      mCallBackPositionInterval = Interval
   Else
      Ret = GRU_ERR_ILLEGALVALUE
   End If
   
SetCallBackPositionInterval_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   SetCallBackPositionInterval = Ret
   Exit Function
   
SetCallBackPositionInterval_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume SetCallBackPositionInterval_Exit
End Function
Public Function GetCallBackPositionInterval(ByRef Interval As Long) As Long

   Const FuncName As String = "GetCallBackPositionInterval"
   
   Dim Ret As Long
   
   On Error GoTo GetCallBackPositionInterval_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   Interval = mCallBackPositionInterval
   
GetCallBackPositionInterval_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GetCallBackPositionInterval = Ret
   Exit Function
   
GetCallBackPositionInterval_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GetCallBackPositionInterval_Exit
End Function

Public Sub Terminate()

   Const FuncName As String = "DSSRecorder_Terminate"
   
   On Error Resume Next
   
   'mSettings.WriteString SettingSection, "LastUsedComport", mComPort
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   If mFileOpened Then
      CloseFile
   End If
   Unload frmRecord
   Set frmRecord = Nothing
Class_Terminate_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit
   Exit Sub
   
Class_Terminate_Err:
   ErrorHandler FuncName, Err
   Resume Class_Terminate_Exit
End Sub
Private Sub frmRecord_DssOnError(ByVal errString As String, ByVal ErrNum As Long)

   Const FuncName As String = "DssOnError"

   ErrorHandler FuncName, ErrNum, errString
End Sub

Private Sub frmRecord_DssOnModeChanged(ByVal mode As Long)

   Const FuncName As String = "DssOnModeChanged"
   
   On Error GoTo DssOnModeChanged_Err
   
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleEntry, CStr(mode)

   Select Case mode
      Case 524, 525
         mState = GRU_STOPPED
      Case 526
         If mRewinding Then
            mState = GRU_REWIND
         Else
            If mFastForwarding Then
               mState = GRU_FORWARD
            Else
               mState = GRU_PLAY
            End If
         End If
      Case 527
         mState = GRU_REC
      'Case 528
      '   mState = GRU_FORWARD  '!!!
      Case 529
         If mRecording Then
            mState = GRU_RECPAUSED
         Else
            mState = GRU_STOPPED
         End If
      Case 530
         mState = GRU_STOPPED
      Case Else
         ErrorHandler "DssOnModeChanged", GRU_ERR_UNKNOWNSTATE
         mState = GRU_UNKNOWN
   End Select
   ReportStateChanged
   
DssOnModeChanged_Exit:
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleExit, CStr(mode)
   Exit Sub
   
DssOnModeChanged_Err:
   ErrorHandler FuncName, Err
   Resume DssOnModeChanged_Exit
End Sub

Private Sub ReportStateChanged()
  
   If mState <> mLastReportedState Then
      mLastReportedState = mState
      RaiseGruEvent GRU_STATECHANGED, mState
   End If
End Sub
Private Sub frmRecord_DssOnPositionChanged(ByVal Position As Long)

   Const FuncName As String = "DssOnPositionChanged"

   On Error GoTo DssOnPositionChanged_Err
   
   'Client.Trace.AddRow Trace_Level_Events, Modulename, FuncName, GruTraceTitleEntry, CStr(Position)
   ReportPosition Position, True
   
DssOnPositionChanged_Exit:
   'Client.Trace.AddRow Trace_Level_Events, Modulename, FuncName, GruTraceTitleExit, CStr(Position)
   Exit Sub
   
DssOnPositionChanged_Err:
   ErrorHandler FuncName, Err
   Resume DssOnPositionChanged_Exit
End Sub
Private Sub ReportPosition(ByVal Position As Long, UseInterval As Boolean)

   Const FuncName As String = "ReportPosition"
   
   If Position < 0 Then
      'GetPos Position
   End If

   On Error GoTo ReportPosition_Err
   
   'Client.Trace.AddRow Trace_Level_Events, Modulename, FuncName, GruTraceTitleEntry, CStr(Position)
   If mCallBackPositionInterval > 0 Then
      If (MyGetTickCount() - mTimeForReportedPosition >= mCallBackPositionInterval) Or Not UseInterval Then
         mTimeForReportedPosition = MyGetTickCount
         If Position < 0 Then
            Position = frmRecord.DssRecorderBase.Position
            Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Position (get)", CStr(Position)
         End If
         If mLastReportedPosition <> Position Then
            mLastReportedPosition = Position
            RaiseGruEvent GRU_POSCHANGE, mLastReportedPosition
         End If
      End If
   End If
   
ReportPosition_Exit:
   'Client.Trace.AddRow Trace_Level_Events, Modulename, FuncName, GruTraceTitleExit, CStr(Position)
   Exit Sub
   
ReportPosition_Err:
   ErrorHandler FuncName, Err
   Resume ReportPosition_Exit
End Sub
Private Sub frmRecord_DssOnRecord(ByVal RecStart As Long, ByVal RecEnd As Long, ByVal FileLength As Long, ByVal PeekPerc As Double)

   Const FuncName As String = "DssOnRecord"

   On Error GoTo DssOnRecord_Err
   If mCallBackPositionInterval > 0 Then
      If MyGetTickCount() - mTimeForReportedPosition >= mCallBackPositionInterval Then
         mTimeForReportedPosition = MyGetTickCount
         mLastReportedPosition = RecEnd
         RaiseGruEvent GRU_POSCHANGE, mLastReportedPosition
      End If
   End If
   RaiseGruEvent GRU_INPUTCHANGE, CLng(PeekPerc)
   
DssOnRecord_Exit:
   'Client.Trace.AddRow Trace_Level_Events, Modulename, FuncName, GruTraceTitleExit, CStr(Position)
   Exit Sub
   
DssOnRecord_Err:
   ErrorHandler FuncName, Err
   Resume DssOnRecord_Exit
End Sub

Private Sub frmRecord_MicStat(Stat As Long)

   Const FuncName As String = "MicStat"
   
   Dim NewButton As Integer

   On Error GoTo MicStat_err
   
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleEntry, CStr(Stat)
   If Stat <> 20 Then
      Select Case Stat
         Case 128
            NewButton = GRU_BUT_PLAYPAUSE
         Case 129
            NewButton = GRU_BUT_REWIND
         Case 130
            NewButton = GRU_BUT_FORWARD
         Case 131
            NewButton = GRU_BUT_PLAY
         Case 134
            NewButton = GRU_BUT_DICT
         Case 140
            NewButton = GRU_BUT_INDEX
         Case 141
            NewButton = GRU_BUT_REC
         Case 142
            NewButton = GRU_BUT_RECPAUSE
         Case 143
            NewButton = GRU_BUT_INSERT
         Case Else
            NewButton = GRU_BUT_UNKNOWN
      End Select
      If NewButton <> mButton Then
         Select Case NewButton
            Case GRU_BUT_PLAYPAUSE
               If Not mRecording Then
                  mButton = NewButton
                  PlayPauseIntern True
               End If
            Case GRU_BUT_REWIND
               mButton = NewButton
               Rewind
            Case GRU_BUT_FORWARD
               mButton = NewButton
               FastForward
            Case GRU_BUT_PLAY
               If Not mRecording Then
                  mButton = NewButton
                  Play
               End If
            Case GRU_BUT_DICT
               mButton = NewButton
               Rec
            Case GRU_BUT_REC
               mButton = NewButton
               RecResume
            Case GRU_BUT_RECPAUSE
               mButton = NewButton
               PlayPauseIntern True
            Case GRU_BUT_INDEX            'To be handled in calling app
               mButton = NewButton
            Case GRU_BUT_INSERT
               mButton = NewButton        'To be handled in calling app
            Case Else
               'ErrorHandler "MicStat", GRU_ERR_UNKNOWNBUTTON
               mButton = GRU_BUT_UNKNOWN
         End Select
         
         RaiseGruEvent GRU_BUTTONPRESS, mButton
      End If
      
   End If
   
MicStat_Exit:
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleExit, CStr(mButton)
   Exit Sub
   
MicStat_err:
   ErrorHandler FuncName, Err
   Resume MicStat_Exit
End Sub
Private Function ErrorHandler(ErrLocation As String, ErrNum As Long, Optional ErrDesc As String = "") As Long

   mErrCode = ErrNum
   mErrDesc = ""
   
   If Len(ErrDesc) = 0 Then
      Select Case mErrCode
         Case GRU_ERR_NOTIMPLEMENTED
            mErrDesc = "Not implemented"
         Case GRU_ERR_UNKNOWNBUTTON
            mErrDesc = "Unknown button"
         Case GRU_ERR_UNKNOWNSTATE
            mErrDesc = "Unknown state"
         Case Else
            mErrDesc = Error$
      End Select
   Else
      mErrDesc = ErrDesc
   End If
   
   ErrorHandler = mErrCode
   
   Client.Trace.AddRow Trace_Level_FatalErrors, ErrLocation, "Error", CStr(ErrNum), ErrDesc
End Function
Public Function LastError(ByRef ErrCode As Long, ByRef ErrDesc As String) As Long

   ErrCode = mErrCode
   ErrDesc = mErrDesc
   
   LastError = 0
End Function

Private Sub RaiseGruEvent(EventType As Gru_Event, Data As Long)

   RaiseEvent GruEvent(EventType, Data)
End Sub

Private Sub frmRecord_MicStatHW(Hw As String)

   Const FuncName As String = "MicStatHW"

   On Error GoTo MicStatHW_Err
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleEntry, Hw
   GetHardWare mHWtype
   RaiseGruEvent GRU_HWCHANGED, mHWtype
   
MicStatHW_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(mHWtype)
   Exit Sub
   
MicStatHW_Err:
   ErrorHandler FuncName, Err
   Resume MicStatHW_Exit
End Sub
'Public Property Get Adapter() As AdapterControl

'   Set Adapter = frmRecord.Adapter
'End Property

