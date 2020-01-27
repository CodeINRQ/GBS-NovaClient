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

