VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SettingSection As String = "Gru_Trace"

Enum TraceLevelType
   Trace_Level_None = 0
   Trace_Level_FatalErrors = 1
   Trace_Level_NonFatalErrors = 2
   Trace_Level_Warning = 4
   Trace_Level_FunctionCalls = 8
   Trace_Level_Events = 16
   Trace_Level_Full = 255
End Enum

Enum TraceOutputType
   Trace_Output_None = 0
   Trace_Output_File = 1
   Trace_Output_ViewForm = 2
   Trace_Output_MsgBox = 4
End Enum

Private mLevel As Integer
Private mOutput As Integer

Private mIniFilePath As String
Private mIniFileSection As String

Private mTraceFilename As String
Private mTraceFileMaxSize As Long
Private mTraceFileCacheRows As Long
Private mTraceFileCreateNew As Boolean

Private mTraceFormAlwaysOnTop As Boolean
Private mTracePosition As String
Private mTraceFormSavePosition As Boolean

Private mCachedRows(100) As String
Private mNumberOfCachedRows As Integer
Private mFileUsed As Boolean
Private mTraceForm As frmTraceOrg
Private mTraceFormUsed As Boolean

Private mSettings As clsSettings
  
Private Sub Class_Terminate()

   On Error Resume Next
   If mTraceFormSavePosition Then
      mSettings.WriteString SettingSection, "FormPosition", WindowSavePositionToString(mTraceForm)
   End If
   mTraceFileCacheRows = 0
   AddRowToFile Trace_Level_Full, "", "", ""
   Unload mTraceForm
   Set mTraceForm = Nothing
End Sub

Public Sub Initialize(Setting As clsSettings)

   Set mSettings = Setting
   
   mLevel = mSettings.GetLong(SettingSection, "Level", Trace_Level_None)
   If mLevel <> Trace_Level_None Then
      mOutput = mSettings.GetLong(SettingSection, "Output", Trace_Output_None)
      
      mTraceFilename = mSettings.GetString(SettingSection, "Filename", App.Path & "\" & App.EXEName & ".log")
      mTraceFileMaxSize = mSettings.GetLong(SettingSection, "FileMaxSize", 0)
      mTraceFileCacheRows = mSettings.GetLong(SettingSection, "FileCacheRows", 50)
      mTraceFileCreateNew = mSettings.GetBool(SettingSection, "FileCreateNew", True)
      
      mTraceFormAlwaysOnTop = mSettings.GetBool(SettingSection, "FormAlwaysOnTop", False)
      mTracePosition = mSettings.GetString(SettingSection, "FormPosition", "")
      mTraceFormSavePosition = mSettings.GetBool(SettingSection, "FormSavePosition", True)
   End If
End Sub
Public Sub AddRow(ByVal Level As TraceLevelType, ByVal Module As String, ByVal Location As String, ByVal Title As String, Optional ByVal Data1 As String, Optional ByVal Data2 As String)

   Dim Data As String
   
   Data = Data1
   Data = Data & " " & Data2
   On Error Resume Next
   
   If (mOutput) = 0 Then Exit Sub
   If (Level And mLevel) = 0 Then Exit Sub
      
   If (mOutput And Trace_Output_File) <> 0 Then
      AddRowToFile Level, Location, Title, Data
   End If
   If (mOutput And Trace_Output_ViewForm) <> 0 Then
      AddRowToForm Level, Location, Title, Data
   End If
   If (mOutput And Trace_Output_MsgBox) <> 0 Then
      AddRowToMsgBox Level, Location, Title, Data
   End If
End Sub
Private Sub AddRowToFile(ByVal Level As TraceLevelType, ByVal Location As String, ByVal Title As String, ByVal Data As String)

   Dim F As Integer
   Dim I As Integer

   On Error Resume Next
   If mNumberOfCachedRows > UBound(mCachedRows) Then Exit Sub
   
   mCachedRows(mNumberOfCachedRows) = Format$(Now, "ddddd ttttt") & vbTab & CStr(Level) & vbTab & Location & vbTab & Title & vbTab & Data
   mNumberOfCachedRows = mNumberOfCachedRows + 1
   
   If mNumberOfCachedRows >= mTraceFileCacheRows Then
      If Not mFileUsed Then
         If mTraceFileCreateNew Then
            Kill mTraceFilename
         End If
         mFileUsed = True
      End If
      If FileLen(mTraceFilename) < mTraceFileMaxSize Then
         F = FreeFile
         Open mTraceFilename For Append As F
         For I = 0 To mNumberOfCachedRows - 1
            Print #F, mCachedRows(I)
         Next I
         Close F
      End If
      mNumberOfCachedRows = 0
   End If
End Sub
Private Sub AddRowToForm(ByVal Level As TraceLevelType, ByVal Location As String, ByVal Title As String, ByVal Data As String)

   Dim Row As String

   On Error Resume Next
   If Not mTraceFormUsed Then
      Set mTraceForm = New frmTraceOrg
      WindowSetPositionFromString mTraceForm, mTracePosition
      If mTraceFormAlwaysOnTop Then
         WindowFloating mTraceForm
      End If
      mTraceForm.Show vbModeless
      mTraceFormUsed = True
   End If
   
   Row = Format$(Now, "ddddd ttttt") & vbTab & CStr(Level) & vbTab & Location & vbTab & Title & vbTab & Data
   mTraceForm.lstTrace.AddItem Row
   mTraceForm.lstTrace.ListIndex = mTraceForm.lstTrace.ListCount - 1
   If mTraceForm.lstTrace.ListCount > 500 Then
      mTraceForm.lstTrace.RemoveItem 0
   End If
End Sub
Private Sub AddRowToMsgBox(ByVal Level As TraceLevelType, ByVal Location As String, ByVal Title As String, ByVal Data As String)

   Dim Row As String
   Dim MsgType As Integer

   On Error Resume Next
   Row = Format$(Now, "ddddd ttttt") & vbLf & CStr(Level) & vbLf & Location & vbLf & Title & vbLf & Data
   If Level < Trace_Level_Warning Then
      MsgType = vbCritical
   Else
      If Level = Trace_Level_Warning Then
         MsgType = vbExclamation
      Else
         MsgType = vbInformation
      End If
   End If
   MsgBox Row, MsgType + vbOKOnly, App.EXEName
End Sub

Public Property Let Level(ByVal NewValue As TraceLevelType)

    mLevel = NewValue
End Property
Public Property Let Output(ByVal NewValue As TraceOutputType)

    mOutput = NewValue
End Property

