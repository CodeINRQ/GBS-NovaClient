VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ExtSystem As String
Public EventName As String

Public Clipboard As String
Public ClipboardDelimit As Integer

Public FilePath As String
Public FileContent As String
Public FileDelimit As Integer

Public Cmd As String
Public CmdParams As String
Public CmdWindowStyle As Integer
Public CmdDelimit As Integer

Public Obj As String

Public CtCmdFile As String
Public CtCmdFileDate As Date

Public WindowModuleName As String
Public WindowUseTopLevel As Boolean
Public WindowIdentControlId As String
Public WindowIdentClassName As String
Public WindowIdentCaption As String
Public WindowIdentParse As String
Public WindowIdentWindowRectangle As String
Public WindowIdentClientRectangle As String

Public IdCheckCompatibility As Boolean

Public IdControl As clsEventWindowControl
Public NameControl As clsEventWindowControl
Public KeywordControl As clsEventWindowControl
Public OrgControl As clsEventWindowControl
Public PriControl As clsEventWindowControl
Public DictTypeControl As clsEventWindowControl
Public TipControl As clsEventWindowControl

Public WindowHndl As Long
Public WindowLastCheckedHndl As Long
Public WindowLastPatid As String
Public WindowLastName As String
Public WindowLastKeyword As String
Public WindowLastOrg As String
Public WindowLastPri As String
Public WindowLastDictType As String


Private Sub Class_Initialize()

   Set IdControl = New clsEventWindowControl
   Set NameControl = New clsEventWindowControl
   Set KeywordControl = New clsEventWindowControl
   Set OrgControl = New clsEventWindowControl
   Set PriControl = New clsEventWindowControl
   Set DictTypeControl = New clsEventWindowControl
   Set TipControl = New clsEventWindowControl
End Sub

Private Sub Class_Terminate()

   Set IdControl = Nothing
   Set NameControl = Nothing
   Set KeywordControl = Nothing
   Set OrgControl = Nothing
   Set PriControl = Nothing
   Set DictTypeControl = Nothing
   Set TipControl = Nothing
End Sub
