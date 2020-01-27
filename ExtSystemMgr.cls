VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtSystemMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mExtSystemArr() As clsExtSystem
Private mExtSystemCount As Integer

Private mExtSystemDictTypeArr() As clsExtSystemDictType
Private mExtSystemDictTypeCount As Integer

Private mExtSystemOrgArr() As clsExtSystemOrg
Private mExtSystemOrgCount As Integer

Private mExtSystemPriorityArr() As clsExtSystemPriority
Private mExtSystemPriorityCount As Integer

Public ExtVariables As Collection

Public Sub AddExtVariable(ByVal ExtSystem As String, ByVal VariableName As String, VariableValue As String)

   Dim Ev As clsExtVariable
   
   ExtSystem = UCase(ExtSystem)
   VariableName = UCase(VariableName)
   
   For Each Ev In ExtVariables
      Ev.VariableValue = VariableValue
      Exit Sub
   Next Ev
   Set Ev = New clsExtVariable
   Ev.ExtSystem = ExtSystem
   Ev.VariableName = VariableName
   Ev.VariableValue = VariableValue
   
   ExtVariables.Add Ev
End Sub
Public Function GetExtVariable(ByVal ExtSystem As String, ByVal VariableName As String, ByRef VariableValue As String) As Boolean

   Dim Ev As clsExtVariable
   
   ExtSystem = UCase(ExtSystem)
   VariableName = UCase(VariableName)
   For Each Ev In Client.ExtSystemMgr.ExtVariables
      If ExtSystem = Ev.ExtSystem And VariableName = Ev.VariableName Then
         VariableValue = Ev.VariableValue
         GetExtVariable = True
         Exit Function
      End If
   Next Ev
   VariableValue = ""
   GetExtVariable = False
End Function

Public Property Get Count() As Integer

   Count = mExtSystemCount
End Property
Public Function ExtSystemName(Idx As Integer) As String

   ExtSystemName = mExtSystemArr(Idx).ExtSystem
End Function
Public Function CheckExtSystem(ExtSys As String) As Boolean

   Dim I As Integer
   
   For I = 0 To mExtSystemCount - 1
      If UCase$(ExtSys) = UCase$(mExtSystemArr(I).ExtSystem) Then
         CheckExtSystem = True
         Exit Function
      End If
   Next I
   CheckExtSystem = False
End Function
Public Function GetIntDictTypeId(ExtSys As String, ExtDictType As String) As Long

   Dim I As Integer
   
   For I = 0 To mExtSystemDictTypeCount - 1
      With mExtSystemDictTypeArr(I)  'not best performance, but easy to read
         If UCase$(ExtSys) = UCase$(.ExtSystem) And UCase$(ExtDictType) = UCase$(.ExtSystemDictType) Then
            GetIntDictTypeId = .IntDictTypeId
            Exit Function
         End If
      End With
   Next I
   GetIntDictTypeId = 0
End Function
Public Function GetIntOrgId(ExtSys As String, ExtOrg As String) As Long

   Dim I As Integer
   
   For I = 0 To mExtSystemOrgCount - 1
      With mExtSystemOrgArr(I)  'not best performance, but easy to read
         If UCase$(ExtSys) = UCase$(.ExtSystem) And UCase$(ExtOrg) = UCase$(.ExtSystemOrg) Then
            GetIntOrgId = .IntOrgId
            Exit Function
         End If
      End With
   Next I
   GetIntOrgId = 0
End Function
Public Function GetIntPriorityId(ExtSys As String, ExtPriority As String) As Long

   Dim I As Integer
   
   For I = 0 To mExtSystemPriorityCount - 1
      With mExtSystemPriorityArr(I)  'not best performance, but easy to read
         If UCase$(ExtSys) = UCase$(.ExtSystem) And UCase$(ExtPriority) = UCase$(.ExtSystemPriority) Then
            GetIntPriorityId = .IntPriorityId
            Exit Function
         End If
      End With
   Next I
   GetIntPriorityId = 0
End Function
Public Function GetExtDictType(ExtSys As String, IntDictTypeId As Long) As String

   Dim I As Integer
   
   For I = 0 To mExtSystemDictTypeCount - 1
      With mExtSystemDictTypeArr(I)  'not best performance, but easy to read
         If UCase$(ExtSys) = UCase$(.ExtSystem) And IntDictTypeId = .IntDictTypeId Then
            GetExtDictType = .ExtSystemDictType
            Exit Function
         End If
      End With
   Next I
   GetExtDictType = ""
End Function
Public Function GetExtOrg(ExtSys As String, IntOrgId As Long) As String

   Dim I As Integer
   
   For I = 0 To mExtSystemOrgCount - 1
      With mExtSystemOrgArr(I)  'not best performance, but easy to read
         If UCase$(ExtSys) = UCase$(.ExtSystem) And IntOrgId = .IntOrgId Then
            GetExtOrg = .ExtSystemOrg
            Exit Function
         End If
      End With
   Next I
   GetExtOrg = ""
End Function
Public Function GetExtPriority(ExtSys As String, IntPriorityId As Long) As String

   Dim I As Integer
   
   For I = 0 To mExtSystemPriorityCount - 1
      With mExtSystemPriorityArr(I)  'not best performance, but easy to read
         If UCase$(ExtSys) = UCase$(.ExtSystem) And IntPriorityId = .IntPriorityId Then
            GetExtPriority = .ExtSystemPriority
            Exit Function
         End If
      End With
   Next I
   GetExtPriority = ""
End Function
Public Sub Init()
   
   mExtSystemCount = 0
   Client.Server.CreateExtSystemList
   Do While Client.Server.ExtSystemGetNext(mExtSystemArr(mExtSystemCount))
      mExtSystemCount = mExtSystemCount + 1
      If mExtSystemCount = UBound(mExtSystemArr) Then
         ReDim Preserve mExtSystemArr(mExtSystemCount + 5)
         ReDim Preserve mExtSystemArr(mExtSystemCount + 5)
      End If
   Loop
   
   mExtSystemDictTypeCount = 0
   Client.Server.CreateExtSystemDictTypeList
   Do While Client.Server.ExtSystemDictTypeGetNext(mExtSystemDictTypeArr(mExtSystemDictTypeCount))
      mExtSystemDictTypeCount = mExtSystemDictTypeCount + 1
      If mExtSystemDictTypeCount = UBound(mExtSystemDictTypeArr) Then
         ReDim Preserve mExtSystemDictTypeArr(mExtSystemDictTypeCount + 5)
         ReDim Preserve mExtSystemDictTypeArr(mExtSystemDictTypeCount + 5)
      End If
   Loop
   
   mExtSystemOrgCount = 0
   Client.Server.CreateExtSystemOrgList
   Do While Client.Server.ExtSystemOrgGetNext(mExtSystemOrgArr(mExtSystemOrgCount))
      mExtSystemOrgCount = mExtSystemOrgCount + 1
      If mExtSystemOrgCount = UBound(mExtSystemOrgArr) Then
         ReDim Preserve mExtSystemOrgArr(mExtSystemOrgCount + 10)
         ReDim Preserve mExtSystemOrgArr(mExtSystemOrgCount + 10)
      End If
   Loop
   
   mExtSystemPriorityCount = 0
   Client.Server.CreateExtSystemPriorityList
   Do While Client.Server.ExtSystemPriorityGetNext(mExtSystemPriorityArr(mExtSystemPriorityCount))
      mExtSystemPriorityCount = mExtSystemPriorityCount + 1
      If mExtSystemPriorityCount = UBound(mExtSystemPriorityArr) Then
         ReDim Preserve mExtSystemPriorityArr(mExtSystemPriorityCount + 10)
         ReDim Preserve mExtSystemPriorityArr(mExtSystemPriorityCount + 10)
      End If
   Loop
   
End Sub

Private Sub Class_Initialize()

   ReDim mExtSystemArr(10)
   ReDim mExtSystemOrgArr(50)
   ReDim mExtSystemDictTypeArr(10)
   ReDim mExtSystemPriorityArr(10)
   
   Set ExtVariables = New Collection
End Sub

Private Sub Class_Terminate()

   Set ExtVariables = Nothing
End Sub
