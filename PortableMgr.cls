VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPortableMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event DeviceChanged()

Private WithEvents DC As DssDigtaConf
Attribute DC.VB_VarHelpID = -1
Private WithEvents DCEx As DssDigtaConfEx
Attribute DCEx.VB_VarHelpID = -1

Private PortableDr As String
Private ImportTempFolder As String

Public Property Set DigtaConf(D As DssDigtaConf)

   Set DC = D
End Property

Public Property Set DigtaConfEx(DEx As DssDigtaConfEx)

   Set DCEx = DEx
End Property

Public Property Get DeviceConnected() As Boolean

   If Len(PortableDr) = 0 Then
      PortableDr = DCEx.GetNextDevice
   End If
   DeviceConnected = Len(PortableDr) > 0
End Property

Public Function MoveAllFilesToImportTempFolder()
   
   Dim Ec As Long
   
   ImportTempFolder = CreateTempFolder("DictImport") & "\"
   
   DC.Open PortableDr, Ec
    
   Dim MoveFiles As String
   Dim MoveOneFile As String
   MoveFiles = Client.SysSettings.ImportMoveFileTypes
   
   Do While Len(MoveFiles) > 0
      MoveOneFile = ConsumeToNextChar(MoveFiles, ";")
      MoveAllFilesToImportTempFolderInt MoveOneFile
   Loop

   DC.Close
End Function

Private Function MoveAllFilesToImportTempFolderInt(FileNameFilter As String)

   Dim DSSPath As String
   DSSPath = DC.DSSPath
   If DSSPath = "" Then Exit Function

   Dim files As DSSHEADERCTRLLib.DssFiles
   Set files = New DSSHEADERCTRLLib.DssFiles
   DC.FilesInPath DSSPath, files, FileNameFilter
    
   If files.Empty = 1 Then Exit Function
    
   DC.Operate1 files, ImportTempFolder, DSSHEADERCTRLLibCtl.OP_MOVE
   
End Function

Public Function GetImportFileName()

   Dim Fn As String
   Fn = Dir(ImportTempFolder & "*.*")
   If Len(Fn) > 0 Then
      GetImportFileName = ImportTempFolder & Fn
   End If
End Function
Public Sub CleanUpAfterImport()

   On Error Resume Next
   RmDir ImportTempFolder
End Sub
Private Sub DC_OnDigtaDeviceChange(ByVal DigtaDriveChar As String, ByVal bArriving As Long, ByVal bIsOurDrive As Long)

   If bArriving > 0 Then
      PortableDr = DigtaDriveChar
   Else
      PortableDr = ""
   End If
   RaiseEvent DeviceChanged
End Sub

