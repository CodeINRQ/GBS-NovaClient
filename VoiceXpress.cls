VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVoiceXpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) _
   As Long

Public Enum vxListeningEnum
   vxListeningUnknown
   vxListeningOff
   vxListeningWakeUp
   vxListeningOn
End Enum

Public Enum vxAppStateEnum
   vxAppStateUnknown = -1
   vxAppStateInit = 0
   vxAppStateMicOn = 1
   vxAppStateMicOff = 2
   vxAppStateNotListening = 8
   vxAppStateSoundNotDefined = 9
   vxAppStateWorking = 11
   vxAppStateQuiting = 12
End Enum

Public Event ChangeListening(NewValue As vxListeningEnum)
Public Event ChangeAppState(NewValue As vxAppStateEnum)

Private WithEvents Vx As SPEECHCENTERLib.SCApplication
Attribute Vx.VB_VarHelpID = -1
Private x As KGuiMgr

Private mLastListening As vxListeningEnum
Private mLastAppState As vxAppStateEnum
Private mVxInstalled As Boolean
Private mVxRunning As Boolean
Private mVxAppState As vxAppStateEnum
Private mAppState As Integer
Private mVxVisible As Boolean
Private mTryToListen As Boolean

Private Sub Class_Initialize()

   mLastListening = vxListeningUnknown
   mLastAppState = vxAppStateUnknown
   
   On Error Resume Next
   Set x = New KGuiMgr
   mVxInstalled = Err = 0
   mVxVisible = VxRunning
End Sub

Private Sub VX_OnChanged(ByVal dispID As Long)

   Dim NewListening As vxListeningEnum
   Dim NewAppState As vxAppStateEnum
   
   'Debug.Print dispID, Vx.Listening, Vx.AppState, Vx.StatusText
   
   mVxVisible = Vx.Visible
   
   Select Case Vx.Listening
       Case 0: NewListening = vxListeningOff
       Case 1: NewListening = vxListeningWakeUp
       Case 2:
               NewListening = vxListeningOn
               mTryToListen = False
       Case Else: NewListening = vxListeningUnknown
   End Select
   
   If NewListening <> mLastListening Then
      mLastListening = NewListening
      RaiseEvent ChangeListening(NewListening)
   End If
   
   Select Case Vx.AppState
      Case 0: NewAppState = vxAppStateInit
      Case 1: NewAppState = vxAppStateMicOn
      Case 2: NewAppState = vxAppStateMicOff
      Case 8: NewAppState = vxAppStateNotListening
      Case 9: NewAppState = vxAppStateSoundNotDefined
      Case 11: NewAppState = vxAppStateWorking
      Case 12: NewAppState = vxAppStateQuiting
      Case Else: NewAppState = vxAppStateUnknown
   End Select
   
   If NewAppState <> mLastAppState Then
      mLastAppState = NewAppState
      RaiseEvent ChangeAppState(NewAppState)
   End If
   
   If NewAppState = vxAppStateQuiting Then
      Set Vx = Nothing
      mVxAppState = vxAppStateUnknown
      mVxVisible = False
   Else
      If mTryToListen And mLastListening <> vxListeningOn Then
         Vx.Listening = 2
      End If
   End If
End Sub
Public Property Get VxInstalled() As Boolean

   VxInstalled = mVxInstalled
End Property
Public Property Get VxRunning() As Boolean

   Dim Res As Long
   
   Res = FindWindow(vbNullString, "Voice Xpress")
   VxRunning = Res <> 0
End Property
Public Property Get VxAppState() As vxAppStateEnum

   VxAppState = mLastAppState
End Property
Public Property Get vxListening() As vxListeningEnum

   vxListening = mLastListening
End Property
Public Property Let vxListening(Value As vxListeningEnum)

   If Not Vx Is Nothing Then
      Select Case Value
         Case vxListeningOff
            Vx.Listening = 0
         Case vxListeningOn
            Vx.Listening = 2
      End Select
   End If
End Property
Public Property Get VxVisible() As Boolean

   VxVisible = mVxVisible
End Property
Public Sub Load()

   If Vx Is Nothing And mVxInstalled Then
      Set Vx = New SPEECHCENTERLib.SCApplication
   End If
End Sub
Public Property Let Activate(Value As Boolean)

   On Error Resume Next
   If Value Then
      Load
      mTryToListen = True
      Vx.Listening = 2
   Else
      mTryToListen = False
      Vx.Listening = 0
   End If
End Property
Public Sub Unload()

   Set Vx = Nothing
End Sub
Public Property Get StatusText() As String

   If Not Vx Is Nothing Then
      StatusText = Vx.StatusText
   End If
End Property
