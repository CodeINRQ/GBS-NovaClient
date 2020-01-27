VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public OrgId               As Long
Public AuthorId            As Long
Public AuthorName          As String
Public TranscriberId       As Long
Public TranscriberName     As String
Public StatusStart         As Integer
Public StatusEnd           As Integer
Public NoPatient           As Boolean
Public Pat                 As clsPat
Public DictTypeId          As Long
Public PriorityId          As Long
Public Txt                 As String

Public RecDateUsed         As Boolean
Public RecStartDate        As Date
Public RecEndDate          As Date

Public TranscribedDateUsed As Boolean
Public TranscribedStartDate As Date
Public TranscribedEndDate As Date

Private Sub Class_Initialize()

   StatusEnd = 255
   Set Pat = New clsPat
   DictTypeId = -1
   PriorityId = -1
End Sub
