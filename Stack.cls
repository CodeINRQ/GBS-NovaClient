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
Private mStackDepth As Integer

Public Sub Push(O As Object)

   Set ObjArr(mStackDepth) = O
   mStackDepth = mStackDepth + 1
   Debug.Print "mStackDepth " & CStr(mStackDepth)
End Sub
Public Sub Pop(O As Object)

   mStackDepth = mStackDepth - 1
   Set O = ObjArr(mStackDepth)
End Sub
Public Property Get StackDepth() As Integer

   mStackDepth = mStackDepth
End Property
