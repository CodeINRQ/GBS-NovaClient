VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mId      As String

Public Property Let Id(Value As String)

   Dim Pos As Integer
   
   Pos = InStr(Value, Chr$(0))
   If Pos > 0 Then
      Value = Left$(Value, Pos - 1)
   End If
   mId = Left$(Value, 50)
End Property
Public Property Get Id() As String

   Id = mId
End Property
