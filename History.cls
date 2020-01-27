VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Rowid               As Long
Private mNumber(12)        As Long
Private mSoundLenSec(12)   As Long

Public Property Get Number(Mth As Integer) As Long

   Number = mNumber(Mth)
End Property
Public Property Let Number(Mth As Integer, Value As Long)

   mNumber(Mth) = Value
End Property
Public Property Get SoundLenSec(Mth As Integer) As Long

   SoundLenSec = mSoundLenSec(Mth)
End Property
Public Property Let SoundLenSec(Mth As Integer, Value As Long)

   mSoundLenSec(Mth) = Value
End Property

