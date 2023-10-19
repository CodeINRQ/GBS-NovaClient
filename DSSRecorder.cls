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

Public Enum Gru_Function
   GRU_FUNC_UNKNOWN = -1    ' -1
   GRU_FUNC_PLAY = 1        ' 0
   GRU_FUNC_PLAYPAUSE       ' 1
   GRU_FUNC_INDEX           ' 2
   GRU_FUNC_DICT            ' 3
   GRU_FUNC_REWIND          ' 4
   GRU_FUNC_FORWARD         ' 5
   GRU_FUNC_REC             ' 6
   GRU_FUNC_RECPAUSE        ' 7
   GRU_FUNC_INSERT          ' 8
   GRU_FUNC_GOTOSTART       ' 9
   GRU_FUNC_GOTOEND         ' 10
   GRU_FUNC_BUTREC          ' 11
   GRU_FUNC_INSERTTOGGLE    ' 12
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
Private mFunction As Gru_Function

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

Public Sub ClearButtonMemory()

   mFunction = 0
End Sub

Public Sub RegisterAndActivateLicens()

   frmRecord.DssRecorderBase.Register
End Sub
Public Function SetMicRecordMode(Start As Integer, Pause As Boolean) As Long

   Const FuncName As String = "SetMicRecordMode"
   
   Dim Ret As Long
   
   On Error GoTo SetMicRecordMode_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, FuncName, GruTraceTitleEntry, CStr(Start)
   
   If Start <> 0 Then
      If mHWtype <> GRU_HW_RECORD Then
         Ret = GRU_ERR_NOHARDWARESUPPORT
      Else
         'If mRecorderInUse Then
         '   Ret = GRU_ERR_ALREADYINUSE
         'Else
            SetRecorderInUse True
            frmRecord.SetMicRecordMode True, mRecMode = GRU_OVERWRITE, Pause
            Ret = 0
         'End If
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
   mComPort = Client.Settings.GetString(SettingSection, "LastUsedComport", "")
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, "10"
   OpenPort mComPort
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, "20"
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
Public Function GetGRUVersion(ByRef version As Long) As Long

   Const FuncName As String = "GetGRUVersion"
   
   Dim Ret As Long

   On Error GoTo GetGRUVersion_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   version = CLng(App.Major) * CLng(1000000) + CLng(App.Minor) * CLng(10000) + CLng(App.Revision)
   
GetGRUVersion_Exit:
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret), CStr(version)
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
      mFunction = GRU_FUNC_UNKNOWN
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
      frmRecord.DssRecorderBase.Stop
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Volume (let)", CStr(mPlayBackVolume)
      frmRecord.DssRecorderBase.Volume = mPlayBackVolume
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Play", CStr(mPlayBackVolume) & ",2"
      frmRecord.DssRecorderBase.Play mPlaySpeed, 2
      mState = GRU_PLAY
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
Play_Exit:
   ReportStateChanged

   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   Play = Ret
   Exit Function
   
Play_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume Play_Exit
End Function

Public Function GotoStart() As Long

   Const FuncName As String = "GotoStart"
   
   Dim Ret As Long
   
   On Error GoTo GotoStart_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   MoveTo 0
   PlayPause
   
GotoStart_Exit:
   ReportStateChanged

   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GotoStart = Ret
   Exit Function
   
GotoStart_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GotoStart_Exit
End Function

Public Function GotoEnd() As Long

   Const FuncName As String = "GotoEnd"
   
   Dim Ret As Long
   Dim L As Long

   On Error GoTo GotoEnd_Err
   
   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry
   
   GetLength L
   MoveTo L
   PlayPause
            
GotoEnd_Exit:
   ReportStateChanged

   Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret)
   GotoEnd = Ret
   Exit Function
   
GotoEnd_Err:
   Ret = ErrorHandler(FuncName, Err)
   Resume GotoEnd_Exit
End Function
Public Function Rec(ButtonDevice As Boolean) As Long

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
            frmRecord.RecPuse = False
            frmRecord.DssRecorderBase.Record Ow
            'SetMicRecordMode True, False
            mDirty = True
            'If mFunction = GRU_FUNC_PLAYPAUSE Or mFunction = GRU_FUNC_RECPAUSE Then   'Removed for SDK1 ver 4
            If Not ButtonDevice Then
               PlayPauseIntern True
            End If
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
         SetMicRecordMode True, False
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
   
   'Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   If mFileOpened Then
      Pos = frmRecord.DssRecorderBase.Position
      'Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Position (get)", CStr(Pos)
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
GetPos_Exit:
   'Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Pos)
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
   
   'Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleEntry, ""
   
   If mFileOpened Then
      Length = frmRecord.DssRecorderBase.Length
      Client.Trace.AddRow Trace_Level_DSSRec, ModuleName, FuncName, "Length", CStr(Length)
   Else
      Ret = GRU_ERR_FILENOTOPENED
   End If
   
GetLength_Exit:
   'Client.Trace.AddRow Trace_Level_FunctionCalls, ModuleName, FuncName, GruTraceTitleExit, CStr(Ret) & " " & CStr(Length)
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
   'Debug.Print mode
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
         frmRecord.RecPuse = False
         mState = GRU_REC
      Case 528
         mState = GRU_FORWARD  '!!!
      Case 529
         If mRecording Then
            mState = GRU_RECPAUSED
            SetMicRecordMode True, True
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

Private Sub Olympus_MicStatEx(micStat As Long, micStatClickType As Long)

   Const FuncName As String = "MicStat"
   
   Dim newFunction As Integer
   Dim QuitRecording As Boolean
   Static WindingDuringRec As Boolean
   Static ButDevice As Boolean
   Static lastMic As Long
   Static SliderPos As Integer
   Dim mic As Long

   On Error GoTo MicStat_err
   
   Debug.Print , , CStr(micStat) & " " & CStr(mState)
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleEntry, CStr(micStat) & "," & CStr(micStatClickType)
   
   mic = micStat * 100 + micStatClickType
   
   If mic = lastMic Then Return
   lastMic = mic
   
   Select Case mic
      Case 12800
         SliderPos = GRU_FUNC_PLAYPAUSE
         newFunction = GRU_FUNC_PLAYPAUSE
      Case 12900
         newFunction = GRU_FUNC_REWIND
      Case 13000
         newFunction = GRU_FUNC_FORWARD
      Case 13100
         SliderPos = GRU_FUNC_PLAY
         newFunction = GRU_FUNC_PLAY
'         Select Case mState
'            Case GRU_RECPAUSED
'               newFunction = GRU_FUNC_REC
'            Case Else
'               newFunction = GRU_FUNC_PLAY
'         End Select
      Case 13400                                      'rec click
         newFunction = GRU_FUNC_DICT
      Case 14000
         newFunction = GRU_FUNC_INDEX
      Case 14100                                      'rec play click
         newFunction = GRU_FUNC_REC
      Case 14200
         newFunction = GRU_FUNC_RECPAUSE
      Case 14300
         mFunction = 0
         newFunction = GRU_FUNC_INSERTTOGGLE
      Case 15000
         If mState = GRU_PLAY Then
            newFunction = GRU_FUNC_PLAYPAUSE
         Else
            newFunction = GRU_FUNC_PLAY
         End If
      'Case 15004
      '   mFunction = 0
      '   newFunction = GRU_FUNC_PLAY
      'Case 15503
      '   SliderPos = GRU_FUNC_REC
      '   newFunction = GRU_FUNC_REC
      Case 15500
         If mState = GRU_REC Then
            newFunction = GRU_FUNC_RECPAUSE
         Else
            newFunction = GRU_FUNC_REC
         End If
      Case 15900   '15504
         SliderPos = GRU_FUNC_PLAYPAUSE
         newFunction = GRU_FUNC_PLAYPAUSE
      Case 15700
         newFunction = GRU_FUNC_GOTOSTART
      Case 15800
         newFunction = GRU_FUNC_GOTOEND
      Case 15900
         newFunction = GRU_FUNC_GOTOEND
      Case 16303
         newFunction = GRU_FUNC_REWIND
      Case 16304
         If SliderPos = GRU_FUNC_PLAY Then
            newFunction = GRU_FUNC_PLAY
         Else
            newFunction = GRU_FUNC_PLAYPAUSE
         End If
      Case 16403
         newFunction = GRU_FUNC_FORWARD
      Case 16404
         If SliderPos = GRU_FUNC_PLAY Then
            newFunction = GRU_FUNC_PLAY
         Else
            newFunction = GRU_FUNC_PLAYPAUSE
         End If
      Case Else
         newFunction = GRU_FUNC_UNKNOWN
   End Select
   
   If newFunction = mFunction Then Exit Sub
   
   If newFunction = GRU_FUNC_UNKNOWN Then
      'Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, "Unknown button", CStr(mic) & "," & CStr(mFunction) & "," & CStr(newFunction) & "," & CStr(mState)
      Exit Sub
   Else
      Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, "Check button", CStr(mic) & "," & CStr(mFunction) & "," & CStr(newFunction) & "," & CStr(mState)
   End If
   
   Select Case newFunction
      Case GRU_FUNC_PLAYPAUSE
         PlayPauseIntern True
      Case GRU_FUNC_REWIND
         Rewind
      Case GRU_FUNC_FORWARD
         FastForward
      Case GRU_FUNC_PLAY
         Play
      Case GRU_FUNC_DICT
         Rec False
      Case GRU_FUNC_BUTREC
         Rec True
      Case GRU_FUNC_REC
         Rec True
      Case GRU_FUNC_RECPAUSE
         PlayPauseIntern True
      Case GRU_FUNC_INDEX              'To be handled in calling app
      Case GRU_FUNC_INSERT
      Case GRU_FUNC_INSERTTOGGLE
      Case GRU_FUNC_GOTOSTART
         GotoStart
      Case GRU_FUNC_GOTOEND
         GotoEnd
      Case Else
         mFunction = GRU_FUNC_UNKNOWN
   End Select
   
   mFunction = newFunction
   
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, "Raise button", CStr(mic) & "," & CStr(mFunction) & "," & CStr(newFunction) & "," & CStr(mState)

   RaiseGruEvent GRU_BUTTONPRESS, mFunction
   
   ReportStateChanged
   
MicStat_Exit:
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleExit, CStr(mFunction)
   Exit Sub
   
MicStat_err:
   ErrorHandler FuncName, Err
   Resume MicStat_Exit
End Sub
Private Sub frmRecord_MicStatEx(micStat As Long, micStatClickType As Long)

   Olympus_MicStatEx micStat, micStatClickType
   Exit Sub

   Const FuncName As String = "MicStat"
   
   Dim newFunction As Integer
   Dim QuitRecording As Boolean
   Static WindingDuringRec As Boolean
   Static ButDevice As Boolean
   Static lastMic As Long
   Dim mic As Long

   On Error GoTo MicStat_err
   
   Debug.Print , , CStr(micStat) & " " & CStr(mState)
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleEntry, CStr(micStat) & "," & CStr(micStatClickType)
   
   mic = micStat * 100 + micStatClickType
   
   If mic = lastMic Then Return
   lastMic = mic
   
   Select Case mic
      Case 12800
         Select Case mState
            Case GRU_REC
               newFunction = GRU_FUNC_RECPAUSE
            Case Else
               newFunction = GRU_FUNC_PLAYPAUSE
         End Select
      Case 12900, 16300
         WindingDuringRec = (mState = GRU_REC) Or (mState = GRU_RECPAUSED)
         newFunction = GRU_FUNC_REWIND
      Case 13000, 16400
         WindingDuringRec = (mState = GRU_REC) Or (mState = GRU_RECPAUSED)
         newFunction = GRU_FUNC_FORWARD
      Case 13100
         newFunction = GRU_FUNC_PLAY
      Case 13400                                      'rec click
         newFunction = GRU_FUNC_DICT
      Case 14000
         newFunction = GRU_FUNC_INDEX
      Case 14100                                      'rec play click
         newFunction = GRU_FUNC_REC
      Case 14200, 1
         newFunction = GRU_FUNC_RECPAUSE
      Case 14300
         newFunction = GRU_FUNC_INSERT
      Case 15000
         mFunction = 0
         Select Case mState
            Case GRU_PLAY
               newFunction = GRU_FUNC_PLAYPAUSE
            Case GRU_REC, GRU_RECPAUSED
               frmRecord.DssRecorderBase.Stop
               mRecording = False
               newFunction = GRU_FUNC_PLAYPAUSE
            Case Else
               newFunction = GRU_FUNC_PLAY
         End Select
      Case 15503
         mFunction = 0
         Select Case mState
            Case GRU_REC
               newFunction = GRU_FUNC_RECPAUSE
            Case GRU_RECPAUSED
               newFunction = GRU_FUNC_REC
            Case Else
               newFunction = GRU_FUNC_REC
         End Select
      Case 15700
         newFunction = GRU_FUNC_GOTOSTART
      Case 15800
         newFunction = GRU_FUNC_GOTOEND
      Case 159000
         Client.Trace.AddRow Trace_Level_Tmp, ModuleName, FuncName, "case 15900+", CStr(mic) & "," & CStr(mFunction) & "," & CStr(newFunction) & "," & CStr(mState)
         ButDevice = True
         mFunction = 0
         Select Case mState
            Case GRU_REC
               newFunction = GRU_FUNC_RECPAUSE
            Case GRU_RECPAUSED
               newFunction = GRU_FUNC_REC
         End Select
         Client.Trace.AddRow Trace_Level_Tmp, ModuleName, FuncName, "case 15900-", CStr(mic) & "," & CStr(mFunction) & "," & CStr(newFunction) & "," & CStr(mState)
         'If lastMic = 12900 Or lastMic = 13000 Then
'            If Not WindingDuringRec Then
'               newFunction = GRU_FUNC_PLAY
'            Else
'               WindingDuringRec = False
'            End If
         'End If
'            Select Case mState
'               Case GRU_REWIND, GRU_FORWARD
'                  If Not WindingDuringRec Then
'                     newFunction = GRU_FUNC_PLAY
'                  Else
'                     WindingDuringRec = False
'                  End If
'            End Select
      Case Else
         newFunction = GRU_FUNC_UNKNOWN
   End Select
   
   If newFunction = mFunction Then Return
      
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, "Check new button", CStr(mic) & "," & CStr(mFunction) & "," & CStr(newFunction) & "," & CStr(mState)
   
   If newFunction = GRU_FUNC_UNKNOWN Then Return
   
   Select Case newFunction
      Case GRU_FUNC_PLAYPAUSE
         If (Not mRecording) Or QuitRecording Then
            mFunction = newFunction
            PlayPauseIntern True
         End If
      Case GRU_FUNC_REWIND
         mFunction = newFunction
         Rewind
      Case GRU_FUNC_FORWARD
         mFunction = newFunction
         FastForward
      Case GRU_FUNC_PLAY
         If Not mRecording Then
            mFunction = newFunction
            Play
         End If
      Case GRU_FUNC_DICT
         mFunction = newFunction
         Rec False
      Case GRU_FUNC_BUTREC
         mFunction = newFunction
         Rec True
      Case GRU_FUNC_REC
         mFunction = newFunction
         RecResume
      Case GRU_FUNC_RECPAUSE
         mFunction = newFunction
         PlayPauseIntern True
      Case GRU_FUNC_INDEX              'To be handled in calling app
         mFunction = newFunction
      Case GRU_FUNC_INSERT
         mFunction = newFunction       'To be handled in calling app
      Case GRU_FUNC_GOTOSTART
         mFunction = newFunction
         GotoStart
      Case GRU_FUNC_GOTOEND
         mFunction = newFunction
         GotoEnd
      Case Else
         mFunction = GRU_FUNC_UNKNOWN
   End Select
   
   RaiseGruEvent GRU_BUTTONPRESS, mFunction
   
   ReportStateChanged
   
MicStat_Exit:
   Client.Trace.AddRow Trace_Level_Events, ModuleName, FuncName, GruTraceTitleExit, CStr(mFunction)
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

