VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CheckDictationForMandatoryData(ByVal OkByFormTest As Boolean, Dict As clsDict, BeforeSave As Boolean) As Boolean
End Function

Public Function CheckPatId(ByVal PatId As String) As Boolean
End Function

Public Function CheckPatname(ByVal PName As String) As Boolean
End Function

Public Function CheckCurrentPatientTextbox(ByVal PatId As String) As Boolean
End Function

Public Function CheckCurrentPatientBeforeUse(ByVal PatId As String) As Boolean
End Function
Public Function RndPatId() As String
End Function

Public Function RndPatName() As String
End Function

Public Function FormatPatId(PatId As String) As String
End Function

Public Function FormatPatIdForStoring(ByVal S As String) As String
End Function

