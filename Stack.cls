VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ObjArr(20) As Object
Private StackDeep As Integer

Public Sub Push(O As Object)

   Set ObjArr(StackDeep) = O
   StackDeep = StackDeep + 1
   Debug.Print "StackDeep " & CStr(StackDeep)
End Sub
Public Sub Pop(O As Object)

   StackDeep = StackDeep - 1
   Set O = ObjArr(StackDeep)
End Sub
