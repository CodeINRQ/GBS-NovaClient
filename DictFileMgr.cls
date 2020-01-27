VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictFileMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function SoundLength(DictFile As clsDictFile) As Long

   Dim SL As String

   On Error Resume Next
   frmMain.DssFileHeaderSimple.Open DictFile.LocalFilenamePlay
   SL = frmMain.DssFileHeaderSimple.Length
   frmMain.DssFileHeaderSimple.Close
   SoundLength = CLng(Left(SL, 2)) * 60 * 60 + CLng(mId(SL, 3, 2)) * 60 + CLng(Right(SL, 2))
End Function
Public Sub KillLocalTempDictationFile(DictFile As clsDictFile)

   If Client.SysSettings.DeleteTemp = 0 Then
      KillFileIgnoreError DictFile.LocalFilenamePlay
      If DictFile.LocalFilenameFormat <> DictFile.LocalFilenamePlay Then
         KillFileIgnoreError DictFile.LocalFilenameFormat
      End If
      DictFile.Clear
   End If
End Sub

Public Sub GetDictFileInfoToSettings(DictFile As clsDictFile, Settings As clsSettings, Section As String)

   DictFile.LocalFilenamePlay = Settings.GetString(Section, "Lfn", "")
   DictFile.LocalFilenameFormat = Settings.GetString(Section, "Lfnf", "")
   DictFile.LocalType = Settings.GetString(Section, "Lt", "")
   DictFile.OrgFileName = Settings.GetString(Section, "Ofn", "")
   DictFile.OrgType = Settings.GetString(Section, "Ot", "")
End Sub
Public Sub WriteDictFileInfoToSettings(DictFile As clsDictFile, Settings As clsSettings, Section As String)

   Settings.WriteString Section, "Lfn", DictFile.LocalFilenamePlay
   Settings.WriteString Section, "Lfnf", DictFile.LocalFilenameFormat
   Settings.WriteString Section, "Lt", DictFile.LocalType
   Settings.WriteString Section, "Ofn", DictFile.OrgFileName
   Settings.WriteString Section, "Ot", DictFile.OrgType
End Sub

Public Sub CreateLocalValues(DictFile As clsDictFile)

   If Len(DictFile.OrgType) = 0 Then
      DictFile.OrgType = "DSS"
   End If
   'If Len(DictFile.LocalFilenameFormat) = 0 Then
      DictFile.LocalFilenameFormat = CreateTempFileName(DictFile.OrgType)
   'End If
   DictFile.LocalFilenamePlay = DictFile.LocalFilenameFormat
End Sub

Public Sub ConvertOrgFormatToPlayFormat(DictFile As clsDictFile)

   Select Case DictFile.OrgType
      Case "DSS", "DS2", "MP3", "WAV"
         DictFile.LocalType = DictFile.OrgType
      Case "M4A"
         DictFile.LocalType = "WAV"
      Case Else
         DictFile.LocalType = ""
   End Select
   
   If DictFile.OrgType = DictFile.LocalType Then
      DictFile.LocalFilenamePlay = DictFile.LocalFilenameFormat
   Else
      DictFile.LocalFilenamePlay = CreateTempFileName(DictFile.LocalType)
   End If
   
   If DictFile.LocalFilenameFormat <> DictFile.LocalFilenamePlay Then
      Select Case DictFile.OrgType
         Case "M4A"
            On Error Resume Next
            ShellAndWait App.Path & "\neroAacDec.exe -if """ & DictFile.LocalFilenameFormat & """ -of """ & DictFile.LocalFilenamePlay & """", vbNormalNoFocus
         Case Else
            FileCopy DictFile.LocalFilenameFormat, DictFile.LocalFilenamePlay
      End Select
   End If
End Sub
Public Function CopyImportFileToTempStorage(Source As String, DictFile As clsDictFile) As Boolean

   DictFile.OrgType = UCase(FileExtensionFromPath(Source))
   DictFile.OrgFileName = FileNameFromPath(Source)
   CreateLocalValues DictFile
   On Error GoTo CopyImportFileToTempStorage_Err
   FileCopy Source, DictFile.LocalFilenameFormat
   ConvertOrgFormatToPlayFormat DictFile
   CopyImportFileToTempStorage = True
   Exit Function
   
CopyImportFileToTempStorage_Err:
   CopyImportFileToTempStorage = False
   Exit Function
End Function

Private Function LastCharPos(s As String, CharSet As String) As Integer

   Dim Pos As Integer
   
   For Pos = Len(s) To 1 Step -1
      If InStr(CharSet, mId(s, Pos, 1)) > 0 Then
         LastCharPos = Pos
         Exit For
      End If
   Next Pos
End Function

Private Function FileNameFromPath(Path As String) As String

   Dim Pos As Integer
   
   Pos = LastCharPos(Path, "\:")
   If Pos < Len(Path) Then
      FileNameFromPath = mId(Path, Pos + 1)
   End If
End Function
Private Function FileExtensionFromPath(Path As String) As String

   Dim Pos As Integer
   
   Pos = LastCharPos(Path, "\:.")
   If Pos < Len(Path) Then
      If mId(Path, Pos, 1) = "." Then
         FileExtensionFromPath = mId(Path, Pos + 1)
      End If
   End If
End Function

