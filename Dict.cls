VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DictId           As Long

Public ExtDictId        As String
Public ExtSystem        As String

Public DictTypeId       As Long
Public DictTypeText     As String

Public PriorityId       As Long
Public PriorityText     As String
Public ExpiryDate       As Date
Public Txt              As String

Public LockedByUserShortName As String
Public LockedByUserLongName  As String
Public LockedByStation  As String
Public LockedTime       As Date

Public StatusId         As Integer
Public StatusText       As String

Public AuthorId               As Long
Public AuthorShortName        As String
Public AuthorLongName         As String

Public TranscriberId          As Long
Public TranscriberOrgId       As Long
Public TranscribedDate        As Date
Public TranscriberShortName   As String
Public TranscriberLongName    As String

Public NoPatient        As Boolean
Public Pat              As clsPat

Public Created          As Date
Public Changed          As Date
Public SoundDeleted     As Date

Public OrgId           As Long
Public OrgText         As String

Public LocalFilename    As String

Public InfoDirty        As Boolean
Public SoundDirty       As Boolean
Public SoundLength      As Long

Public SoundReadOnly    As Boolean
Public TextReadOnly     As Boolean
Public TimeStamp        As Double

Private Sub Class_Initialize()

   Set Pat = New clsPat
End Sub

Private Sub Class_Terminate()

   'no action
End Sub
Public Property Get LoggData() As String

   LoggData = Pat.PatIdFormatted & "/" & DictId
End Property

