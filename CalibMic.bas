Attribute VB_Name = "modCalibMic"
Option Explicit

Private Declare Function VCStoreIni _
    Lib "Helper.dll" _
    Alias "_VolumeControl_StoreIni@4" (ByVal IniFileName As String) As Long

Private Declare Function VCPrepareSettingsForPlayer _
    Lib "Helper.dll" _
    Alias "_VolumeControl_PrepareSettingsForPlayer@8" (ByVal bWithRecorder As Long, _
                                                       ByVal IniFileName As String) As Long

Private Declare Function VCResetSettingsForPlayer _
    Lib "Helper.dll" _
    Alias "_VolumeControl_ResetSettingsForPlayer@0" () As Long
    
Private Declare Function VCBeginRecord _
    Lib "Helper.dll" _
    Alias "_VolumeControl_OnBeginRecord@0" () As Long

Private Declare Function VCEndRecord _
    Lib "Helper.dll" _
    Alias "_VolumeControl_OnEndRecord@0" () As Long

Private Declare Function VCShowRecordSettingsDialog _
    Lib "Helper.dll" _
    Alias "_VolumeControl_ShowRecordSettingsDialog@0" () As Long

Private Declare Function VCUnShowRecordSettingsDialog _
    Lib "Helper.dll" _
    Alias "_VolumeControl_UnShowRecordSettingsDialog@0" () As Long

Private FileForSettings As String

Public Function RestoreCalibration() As Boolean

   Dim AudioSettings As String
   Dim Ver As String
   Dim IsCalibStoredInDb As Boolean
   Dim OldIniFile As String
   
   'IsCalibStoredInDb = Client.Server.ReadStationData("Audio", "Calib", "", Ver) = "Y"
   
   'AudioSettings = Client.Server.ReadStationData("Audio", "Settings", "", Ver)
   'If Len(AudioSettings) > 0 Then
   '   FileForSettings = WriteStringToTempFile(AudioSettings)
   '   VCPrepareSettingsForPlayer 1, FileForSettings
   '   KillFileIgnoreError FileForSettings
   'End If
   'RestoreCalibration = IsCalibStoredInDb
End Function
Public Sub StartCalibration()

   'Load frmCalibMic
End Sub
Public Sub SaveCalibration()

   'SaveAudioSettings
   'Client.Server.WriteStationData "Audio", "Calib", "Y"
End Sub
Public Sub RestoreAudioSettings()

   'VCResetSettingsForPlayer
End Sub
Private Sub SaveAudioSettings()

   Dim AudioSettings As String
   
   FileForSettings = CreateTempFileName("tmp")
   VCStoreIni FileForSettings
   AudioSettings = ReadStringFromTempFile(FileForSettings)
   Client.Server.WriteStationData "Audio", "Settings", AudioSettings
   KillFileIgnoreError FileForSettings
End Sub
Private Function CheckIfThereIsMicCalibrationInIniFile(Fn As String) As Boolean

   Dim ReadFile As Integer
   Dim s As String
   Dim Res As Boolean
   
   On Error Resume Next
   ReadFile = FreeFile
   Open Fn For Input Access Read As ReadFile
   If Err.Number <> 0 Then
      Exit Function
   End If
   Do While Not EOF(ReadFile)
      Line Input #ReadFile, s
      If InStr(LCase$(s), "mixer") > 0 Then
         CheckIfThereIsMicCalibrationInIniFile = True
         Exit Do
      End If
   Loop
   Close ReadFile
End Function
