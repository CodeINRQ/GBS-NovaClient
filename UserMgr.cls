VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mUsersArr(1000) As clsUser
Private mUsersCount As Integer


Public Sub Init()

   mUsersCount = 0
   Client.Server.CreateUserList
   Do While Client.Server.UserListGetNext(mUsersArr(mUsersCount))
      mUsersCount = mUsersCount + 1
   Loop
End Sub

Public Function IndexFromId(Id As Long) As Integer

   Dim I As Integer
 
   For I = 0 To mUsersCount - 1
      If mUsersArr(I).UserId = Id Then
         IndexFromId = I
         Exit For
      End If
   Next I
End Function

Public Function IdFromIndex(Index As Integer) As Long

   IdFromIndex = mUsersArr(Index).UserId
End Function

Public Property Get Count() As Integer

   Count = mUsersCount
End Property
Public Function GetUserFromId(ByRef Usr As clsUser, UserId As Long) As Boolean

   Dim I As Integer
   
   For I = 0 To mUsersCount - 1
      If mUsersArr(I).UserId = UserId Then
         Set Usr = mUsersArr(I)
         GetUserFromId = True
         Exit For
      End If
   Next
End Function
Public Sub GetUserFromIndex(ByRef Usr As clsUser, Index As Integer)

   If Index < mUsersCount Then
      Set Usr = mUsersArr(Index)
   Else
      Set Usr = Nothing
   End If
End Sub
Public Sub GetUserFromLoginName(ByRef Usr As clsUser, LoginName As String)

   Dim I As Integer
   
   For I = 0 To mUsersCount - 1
      If mUsersArr(I).LoginName = UCase$(LoginName) Then
         Set Usr = mUsersArr(I)
         Exit For
      End If
   Next
End Sub
Public Sub SaveUser(Usr As clsUser)

   On Error GoTo SaveUser_Err
   Client.Server.SaveUser Usr
   Client.LoggMgr.Insert 1320107, LoggLevel_UserAdmin, 0, Usr.LoggData
   mUsersCount = 0
   Init
   Exit Sub
   
SaveUser_Err:
   ErrorHandle "1310101", Err, 1310101, "Användaren kan inte sparas", False
   Resume Next
End Sub
Public Sub DeleteUser(Usr As clsUser)

   On Error GoTo DeleteUser_Err
   Client.Server.DeleteUser Usr
   Client.LoggMgr.Insert 1320118, LoggLevel_UserAdmin, 0, Usr.LoggData
   mUsersCount = 0
   Init
   Exit Sub
   
DeleteUser_Err:
   ErrorHandle "1310102", Err, 1310102, "Användaren kan inte raderas", False
   Resume Next
End Sub
Private Function Validate(LoginName As String, Password As String, NewPassword As String) As Integer

   Dim LId As String
   Dim FN As String
   Dim Id As Long

   Validate = Client.Server.UserLogin(Client.User, LoginName, Password, NewPassword)
End Function
Public Function LoginUser() As Integer

   Dim LoginForm As frmLogin
   
   If Client.Server.NeedLoginInfo() Then
   
      If Len(StartUpUserLoginName) = 0 Then
         Set LoginForm = New frmLogin
         CenterForm LoginForm, frmMain
         
         LoginForm.Show vbModal
         If LoginForm.OkClicked Then
            If LoginForm.ChangePassword Then
               LoginUser = Validate(LoginForm.txtUserName, LoginForm.txtPassword, LoginForm.txtNewPassword)
            Else
               LoginUser = Validate(LoginForm.txtUserName, LoginForm.txtPassword, "")
            End If
         Else
            LoginUser = 100
         End If
         
         Unload LoginForm
         Set LoginForm = Nothing
      Else
         LoginUser = Validate(StartUpUserLoginName, StartUpPassword, "")
      End If
   Else
      LoginUser = Validate("", "", "")
   End If
End Function
