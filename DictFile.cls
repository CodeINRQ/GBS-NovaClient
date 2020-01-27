VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public LocalFilenamePlay   As String
Public LocalFilenameFormat As String
Public LocalType           As String
Public OrgType             As String
Public OrgFileName         As String


Public Sub Clear()

   LocalFilenamePlay = ""
   LocalFilenameFormat = ""
   LocalType = ""
   OrgType = ""
   OrgFileName = ""
End Sub

Public Property Get IsSoundData() As Boolean

   IsSoundData = Len(LocalFilenameFormat) > 0 Or Len(LocalFilenamePlay) > 0
End Property
Public Property Get IsSoundToPlay() As Boolean

   IsSoundToPlay = Len(LocalFilenamePlay) > 0
End Property

