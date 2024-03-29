VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSysSettings As clsStringStore
Private mUseAuthorsSign As Boolean

Private mDeleteSound As Integer
Private mDeleteDictInfo As Integer
Private mDeleteHistory As Integer
Public DeleteLogg As Integer
Public DeleteTemp As Integer

Public DictListShowDictId As Boolean
Public DictListUpdateK As Long
Public DictListUpdateM As Long
Public DictListUpdateMax As Long
Public DictListUpdateValues As Long
Public DictListCurrStatusStart As Long
Public DictListCurrStatusEnd As Long

Public VoiceExpressActive As Boolean
Public VoiceExpressShowTab As Boolean
Public VoiceExpressShowInToolbar As Boolean
Public VoiceExpressActivateOnIndexButton As Boolean

Public DictInfoDefaultDictTypeId As Long
Public DictInfoKeepDictTypeNoDef As Boolean
Public DictInfoKeepDictTypeAlways As Boolean
Public DictInfoDefaultPriorityId As Long
Public DictInfoMandatoryPatId As Boolean
Public DictInfoMandatoryPatIdCentury As Boolean
Public DictInfoAlfaInPatid As Boolean
Public DictInfoMandatoryPatName As Boolean
Public DictInfoPatIdChecksum As Boolean
Public DictInfoUseNoPat As Boolean
Public DictInfoUseKeyWords As Boolean

Public LoginMaxRetries As Long
Public LoginLockoutMinutes As Long
Public LoginPasswordMinLength As Long
Public LoginPasswordCaseSensitive As Boolean
Public LoginAllowChangePassword As Boolean

Public PlayerWarningLowRecLevel As Long
Public PlayerWindingSpeed As Long
Public PlayerHWcheckfreq As Integer
Public PlayerCaption As String
Public PlayerForceMicCalib As Boolean
Public PlayerShowEditButtons As Boolean
Public PlayerShowOnTop As Boolean
Public PlayerShowSmallerWindow As Boolean
Public PlayerIndexButtonAsInsertRec As Boolean
Public PlayerAutoOverwrite As Boolean
Public PlayerIndexButtonAsCloseDict As Boolean
Public PlayerAllowInsertFromMic As Boolean
Public PlayerAllowInsertFromGUI As Boolean
Public PlayerKeyPlay As Integer
Public PlayerKeyPause As Integer
Public PlayerKeyStop As Integer
Public PlayerKeyStart As Integer
Public PlayerKeyRewind As Integer
Public PlayerKeyForward As Integer
Public PlayerKeyEnd As Integer
Public PlayerKeyRec As Integer
Public PlayerKeyVolumeUp As Integer
Public PlayerKeyVolumeDown As Integer
Public PlayerKeySpeedUp As Integer
Public PlayerKeySpeedDown As Integer
Public PlayerKeyClose1 As Integer
Public PlayerKeyClose2 As Integer
Public PlayerKeyClose3 As Integer
Public PlayerKeyCloseX As Integer
Public PlayerKeyClose1Alt As Integer
Public PlayerKeyClose2Alt As Integer
Public PlayerKeyClose3Alt As Integer
Public PlayerKeyCloseXAlt As Integer
Public PlayerKeyEscape As Integer
Public PlayerKeyEscapeAlt As Integer
Public PlayerKeyWindowSize As Integer
Public PlayerCheckMandatoryBeforeRecording As Boolean

Public DemoAllowGenerate As Boolean
Public DemoShowTab As Boolean

Public UserAllowChangeHome As Boolean

Public ClientStopLoginMsg As String
Public ClientForceLogoffMsg As String
Public ClientMinVersion As String
Public ClientLoadTimeLimit As Long

Public SoundUseImageField As Boolean
Public SoundChunkSize As Long
Public SoundMaxSizeToCheckIn As Long

Public CultureDefaultLanguage As String
Public CultureAllowChange As Boolean
Public CultureList As String

Public LoggEnable As Boolean
Public LoggFilter As String

Public DictationAuditEnable As Boolean

Public ImportAllowMenu As Boolean
Public ImportAllowTool As Boolean
Public ImportDSSDrives As String
Public ImportDefaultId As String
Public ImportDefaultName As String
Public ImportDefaultKeyWord As String
Public ImportDefaultNote As String
Public ImportDeleteOrgFile As Boolean
Public ImportMoveFileTypes As String
Public ImportDirectToRecorded As Boolean
Public ImportWithDialog As Boolean

Public ExportAllowMenu As Boolean

Public IndicatorActive As Boolean
Public IndicatorStyle As String
Public IndicatorBgColor As Long
Public IndicatorLightColor As Long
Public IndicatorLightTime As Long
Public IndicatorDarkTime As Long
Public IndicatorMaxBlink As Long


Public Property Get Store() As clsStringStore

   Set Store = mSysSettings
End Property
Public Property Set Store(Store As clsStringStore)

   Set mSysSettings = Store
End Property
Private Sub Class_Initialize()

   Set mSysSettings = New clsStringStore
End Sub
Private Sub Class_Terminate()

   Set mSysSettings = Nothing
End Sub
Public Sub Init(Ty As String)

   Dim Se As String
   Dim Ke As String
   Dim Va As String

   mSysSettings.EmptyStore

   Client.Server.CreateSysSettingsList Ty
   Do While Client.Server.SysSettingsListGetNext(Se, Ke, Va)
      mSysSettings.AddString Se, Ke, Va, False
   Loop
   
   mUseAuthorsSign = mSysSettings.GetBool("WorkFlow", "UseAuthorsSign", False)
   
   mDeleteSound = mSysSettings.GetLong("Delete", "Sound", -1)
   mDeleteDictInfo = mSysSettings.GetLong("Delete", "DictInfo", -1)
   mDeleteHistory = mSysSettings.GetLong("Delete", "History", -1)
   DeleteLogg = mSysSettings.GetLong("Delete", "Logg", 30)
   DeleteTemp = mSysSettings.GetLong("Delete", "Temp", 0)
   
   DictListShowDictId = mSysSettings.GetBool("DictList", "ShowDictId", False)
   DictListUpdateK = mSysSettings.GetLong("DictList", "UpdateK", 10)
   DictListUpdateM = mSysSettings.GetLong("DictList", "UpdateM", 2000)
   DictListUpdateMax = mSysSettings.GetLong("DictList", "UpdateMax", 100000)
   DictListUpdateValues = mSysSettings.GetLong("DictList", "UpdateValues", 5)
   DictListCurrStatusStart = mSysSettings.GetLong("DictList", "CurrStatusStart", 20)
   DictListCurrStatusEnd = mSysSettings.GetLong("DictList", "CurrStatusEnd", 60)
   
   VoiceExpressActive = mSysSettings.GetBool("VoiceExpress", "Active", False)
   VoiceExpressShowTab = mSysSettings.GetBool("VoiceExpress", "ShowTab", False)
   VoiceExpressShowInToolbar = mSysSettings.GetBool("VoiceExpress", "ShowInToolbar", False)
   VoiceExpressActivateOnIndexButton = mSysSettings.GetBool("VoiceExpress", "ActivateOnIndexButton", False)
   
   DictInfoDefaultPriorityId = mSysSettings.GetLong("DictInfo", "DefaultPriorityId", 3)
   DictInfoDefaultDictTypeId = mSysSettings.GetLong("DictInfo", "DefaultDictTypeId", 3)
   DictInfoKeepDictTypeNoDef = mSysSettings.GetBool("DictInfo", "KeepDictTypeNoDef", False)
   DictInfoKeepDictTypeAlways = mSysSettings.GetBool("DictInfo", "KeepDictTypeAlways", False)
   DictInfoMandatoryPatId = mSysSettings.GetBool("DictInfo", "MandatoryPatId", True)
   DictInfoMandatoryPatIdCentury = mSysSettings.GetBool("DictInfo", "MandatoryPatIdCentury", False)
   DictInfoAlfaInPatid = mSysSettings.GetBool("DictInfo", "AlfaInPatid", True)
   DictInfoMandatoryPatName = mSysSettings.GetBool("DictInfo", "MandatoryPatName", False)
   DictInfoPatIdChecksum = mSysSettings.GetBool("DictInfo", "PatIdChecksum", False)
   DictInfoUseNoPat = mSysSettings.GetBool("DictInfo", "UseNoPat", True)
   DictInfoUseKeyWords = mSysSettings.GetBool("DictInfo", "UseKeyWords", True)
   
   LoginMaxRetries = mSysSettings.GetLong("Login", "MaxRetries", 5)
   LoginLockoutMinutes = mSysSettings.GetLong("Login", "LockoutMinutes", 5)
   LoginPasswordMinLength = mSysSettings.GetLong("Login", "PasswordMinLength", 5)
   LoginPasswordCaseSensitive = mSysSettings.GetBool("Login", "PasswordCaseSensitive", True)
   LoginAllowChangePassword = mSysSettings.GetBool("Login", "AllowChangePassword", True)
   
   ClientStopLoginMsg = mSysSettings.GetString("Client", "StopLoginMsg", "")
   ClientForceLogoffMsg = mSysSettings.GetString("Client", "ForceLogoffMsg", "")
   ClientMinVersion = mSysSettings.GetString("Client", "MinVersion", "")
   ClientLoadTimeLimit = mSysSettings.GetLong("Client", "LoadTimeLimit", 12)
   
   PlayerWarningLowRecLevel = mSysSettings.GetLong("Player", "WarningLowRecLevel", 5)
   PlayerWindingSpeed = mSysSettings.GetLong("Player", "WindingSpeed", 6000)
   PlayerHWcheckfreq = mSysSettings.GetLong("Player", "HWcheckfreq", 10)
   PlayerCaption = mSysSettings.GetString("Player", "Caption", "")
   PlayerForceMicCalib = mSysSettings.GetBool("Player", "ForceMicCalib", True)
   PlayerShowEditButtons = mSysSettings.GetBool("Player", "ShowEditButtons", True)
   PlayerShowOnTop = mSysSettings.GetBool("Player", "ShowOnTop", True)
   PlayerShowSmallerWindow = mSysSettings.GetBool("Player", "ShowSmallerWindow", True)
   PlayerIndexButtonAsCloseDict = mSysSettings.GetBool("Player", "CloseDict", True)
   PlayerAllowInsertFromMic = mSysSettings.GetBool("Player", "AllowInsertFromMic", True)
   PlayerAllowInsertFromGUI = mSysSettings.GetBool("Player", "AllowInsertFromGUI", True)
   PlayerAutoOverwrite = mSysSettings.GetBool("Player", "AutoOverwrite", True)
   PlayerKeyPlay = mSysSettings.GetLong("Player", "KeyPlay", vbKeyF2)
   PlayerKeyPause = mSysSettings.GetLong("Player", "KeyPause", vbKeyF3)
   PlayerKeyStop = mSysSettings.GetLong("Player", "KeyStop", vbKeyF4)
   PlayerKeyStart = mSysSettings.GetLong("Player", "KeyStart", vbKeyF5)
   PlayerKeyRewind = mSysSettings.GetLong("Player", "KeyRewind", vbKeyF6)
   PlayerKeyForward = mSysSettings.GetLong("Player", "KeyForward", vbKeyF7)
   PlayerKeyEnd = mSysSettings.GetLong("Player", "KeyEnd", vbKeyF8)
   PlayerKeyRec = mSysSettings.GetLong("Player", "KeyRec", vbKeyF9)
   PlayerKeyVolumeUp = mSysSettings.GetLong("Player", "KeyVolumeUp", vbKeyAdd + 1 * 256)
   PlayerKeyVolumeDown = mSysSettings.GetLong("Player", "KeyVolumeDown", vbKeySubtract + 1 * 256)
   PlayerKeySpeedUp = mSysSettings.GetLong("Player", "KeySpeedUp", vbKeyAdd + 2 * 256)
   PlayerKeySpeedDown = mSysSettings.GetLong("Player", "KeySpeedDown", vbKeySubtract + 2 * 256)
   PlayerKeyClose1 = mSysSettings.GetLong("Player", "KeyClose1", vbKeyF10)
   PlayerKeyClose2 = mSysSettings.GetLong("Player", "KeyClose2", vbKeyF11)
   PlayerKeyClose3 = mSysSettings.GetLong("Player", "KeyClose3", vbKeyF12)
   PlayerKeyCloseX = mSysSettings.GetLong("Player", "KeyCloseX", 0)
   PlayerKeyEscape = mSysSettings.GetLong("Player", "KeyEscape", 27)
   PlayerKeyClose1Alt = mSysSettings.GetLong("Player", "KeyClose1Alt", 1073)
   PlayerKeyClose2Alt = mSysSettings.GetLong("Player", "KeyClose2Alt", 1074)
   PlayerKeyClose3Alt = mSysSettings.GetLong("Player", "KeyClose3Alt", 1075)
   PlayerKeyCloseXAlt = mSysSettings.GetLong("Player", "KeyCloseXAlt", 1089)
   PlayerKeyEscapeAlt = mSysSettings.GetLong("Player", "KeyEscapeAlt", 0)
   PlayerKeyWindowSize = mSysSettings.GetLong("Player", "KeyWindowSize", 1114)   'Alt-Z
   PlayerCheckMandatoryBeforeRecording = mSysSettings.GetBool("Player", "CheckMandatoryBeforeRecording", False)  'Alt-Z
   
   DemoAllowGenerate = mSysSettings.GetBool("Demo", "AllowGenerate", False)
   DemoShowTab = mSysSettings.GetBool("Demo", "ShowTab", False)
   
   UserAllowChangeHome = mSysSettings.GetBool("User", "AllowChangeHome", False)
   
   SoundUseImageField = mSysSettings.GetBool("Sound", "UseImageField", True)
   SoundChunkSize = mSysSettings.GetLong("Sound", "ChunkSize", 8192)
   SoundMaxSizeToCheckIn = mSysSettings.GetLong("Sound", "MaxSizeToCheckIn", 0)
   
   CultureDefaultLanguage = mSysSettings.GetString("Culture", "DefaultLanguage", "SE")
   CultureAllowChange = mSysSettings.GetBool("Culture", "AllowChange", True)
   CultureList = mSysSettings.GetString("Culture", "DefaultList", "SE;EN;DK;NO")
   
   LoggEnable = mSysSettings.GetBool("Logg", "Enable", True)
   LoggFilter = mSysSettings.GetString("Logg", "Filter", "")
   
   DictationAuditEnable = mSysSettings.GetBool("DictationAudit", "Enable", True)
   
   ImportAllowMenu = mSysSettings.GetBool("Import", "AllowMenu", True)
   ImportAllowTool = mSysSettings.GetBool("Import", "AllowTool", True)
   ImportDSSDrives = mSysSettings.GetString("Import", "DSSDrives", "")
   ImportDefaultId = mSysSettings.GetString("Import", "DefaultId", "")
   ImportDefaultName = mSysSettings.GetString("Import", "DefaultName", "")
   ImportDefaultKeyWord = mSysSettings.GetString("Import", "DefaultKeyWord", "")
   ImportDefaultNote = mSysSettings.GetString("Import", "DefaultNote", "")
   ImportDeleteOrgFile = mSysSettings.GetBool("Import", "DeleteOrgFile", True)
   ImportMoveFileTypes = mSysSettings.GetString("Import", "MoveFiles", "*.DSS;*.DS2;*.MP3;*.WAV")
   ImportDirectToRecorded = mSysSettings.GetBool("Import", "DirectToRecorded", False)
   ImportWithDialog = mSysSettings.GetBool("Import", "WithDialog", False)


   ExportAllowMenu = mSysSettings.GetBool("Export", "AllowMenu", True)

   IndicatorActive = mSysSettings.GetBool("Indicator", "Active", True)
   IndicatorStyle = mSysSettings.GetString("Indicator", "Style", "")
   IndicatorBgColor = mSysSettings.GetLong("Indicator", "BgColor", 8454143)
   IndicatorLightColor = mSysSettings.GetLong("Indicator", "LightColor", 49152)
   IndicatorLightTime = mSysSettings.GetLong("Indicator", "LightTime", 2)
   IndicatorDarkTime = mSysSettings.GetLong("Indicator", "DarkTime", 20)
   IndicatorMaxBlink = mSysSettings.GetLong("Indicator", "MaxBlink", -1)
   
End Sub
Public Sub Save(Ty As String)

   Dim Se As String
   Dim Ke As String
   Dim Va As String
   
   mSysSettings.Filter = ""
   Do While mSysSettings.GetNextFromFilter(Se, Ke, Va)
      Client.Server.WriteSysSettings Ty, Se, Ke, Va
   Loop
   Client.LoggMgr.Insert 1320116, LoggLevel_SysAdmin, 0, ""
End Sub
Public Property Get UseAuthorsSign() As Boolean

   UseAuthorsSign = mUseAuthorsSign
End Property
Public Property Get DeleteSound() As Long

   DeleteSound = mDeleteSound
End Property
Public Property Get DeleteDictInfo() As Long

   DeleteDictInfo = mDeleteDictInfo
End Property
Public Property Get DeleteHistory() As Long

   DeleteHistory = mDeleteHistory
End Property
