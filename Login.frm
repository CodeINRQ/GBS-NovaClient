VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CareTalk Login"
   ClientHeight    =   2550
   ClientLeft      =   2835
   ClientTop       =   3360
   ClientWidth     =   5610
   HelpContextID   =   1010000
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1506.623
   ScaleMode       =   0  'User
   ScaleWidth      =   5267.487
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1010100"
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":1CF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ImageList1"
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   1560
      Width           =   2325
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1080
      Width           =   2325
   End
   Begin VB.CheckBox chkChangePassword 
      Caption         =   "&Byt lösenord"
      Height          =   255
      Left            =   81
      TabIndex        =   8
      Tag             =   "1010105"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2400
      TabIndex        =   4
      Tag             =   "1010106"
      Top             =   2040
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Avbryt"
      Height          =   390
      Left            =   3600
      TabIndex        =   5
      Tag             =   "1010107"
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label lblMarker 
      Height          =   255
      Left            =   105
      TabIndex        =   15
      Top             =   2040
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.Label lblConfirmPassword 
      Caption         =   "&Bekräfta lösenord:"
      Height          =   270
      Left            =   105
      TabIndex        =   14
      Tag             =   "1010104"
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label lblNewPassword 
      Caption         =   "&Nytt lösenord:"
      Height          =   270
      Left            =   105
      TabIndex        =   13
      Tag             =   "1010103"
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label lblNewPasswordMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4710
      TabIndex        =   12
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label lblConfirmPasswordMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4710
      TabIndex        =   11
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label lblPasswordMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4710
      TabIndex        =   7
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblUserIdMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4710
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Användarid:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "1010101"
      Top             =   150
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Lösenord:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Tag             =   "1010102"
      Top             =   600
      Width           =   1680
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OkClicked As Boolean
Public ChangePassword As Boolean

Private Sub chkChangePassword_Click()

   ChangePassword = chkChangePassword.Value = vbChecked
   If ChangePassword Then
      ExpandWindow
   Else
      CollapsWindow
   End If
   SetEnabled
End Sub

Private Sub cmdCancel_Click()
    
    OkClicked = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
   OkClicked = True
   Me.Hide
End Sub

Private Sub Form_Activate()

   SetWindowTopMostAndForeground Me
End Sub

Private Sub Form_Load()

   CenterAndTranslateForm Me, frmMain

   If Client.SysSettings.CultureAllowChange Then
      ImageCombo1.ComboItems.Add 1, "SE", , 1
      ImageCombo1.ComboItems.Add 2, "EN", , 2
      ImageCombo1.ComboItems.Add 3, "DK", , 3
      ImageCombo1.ComboItems.Add 4, "NO", , 4
      ImageCombo1.ComboItems.Add 5, "FI", , 5
      ImageCombo1.ComboItems.Add 6, "DE", , 6
      ImageCombo1.ComboItems.Add 7, "FR", , 7
      ImageCombo1.SelectedItem = ImageCombo1.ComboItems(Client.SysSettings.CultureDefaultLanguage)
      ImageCombo1.Visible = True
   Else
      ImageCombo1.Visible = False
   End If
   ChangePassword = False
   CollapsWindow
   chkChangePassword.Visible = Client.SysSettings.LoginAllowChangePassword
End Sub

Private Sub CollapsWindow()

   Dim T As Integer

   lblNewPassword.Visible = False
   txtNewPassword.Visible = False
   lblNewPasswordMissing.Visible = False
   
   lblConfirmPassword.Visible = False
   txtConfirmPassword.Visible = False
   lblConfirmPasswordMissing.Visible = False
   
   T = lblNewPassword.Top
   chkChangePassword.Top = T
   cmdOK.Top = T
   cmdCancel.Top = T
   
   Me.Height = 2055
End Sub
Private Sub ExpandWindow()

   Dim T As Integer

   lblNewPassword.Visible = True
   txtNewPassword.Visible = True
   lblNewPasswordMissing.Visible = True
   
   lblConfirmPassword.Visible = True
   txtConfirmPassword.Visible = True
   lblConfirmPasswordMissing.Visible = True
   
   T = lblMarker.Top
   chkChangePassword.Top = T
   cmdOK.Top = T
   cmdCancel.Top = T
   
   Me.Height = 2970
End Sub
Private Sub SetEnabled()

   Dim Ok As Boolean
   
   Ok = True
   If Len(txtUserName.Text) = 0 Then
      lblUserIdMissing.Visible = True
      Ok = False
   Else
      lblUserIdMissing.Visible = False
   End If
   If Len(txtPassword.Text) = 0 Then
      lblPasswordMissing.Visible = True
      Ok = False
   Else
      lblPasswordMissing.Visible = False
   End If
   
   If ChangePassword Then
      If Len(txtNewPassword.Text) < Client.SysSettings.LoginPasswordMinLength Then
         lblNewPasswordMissing.Visible = True
         Ok = False
      Else
         lblNewPasswordMissing.Visible = False
      End If
      If txtNewPassword.Text <> txtConfirmPassword.Text Or Len(txtConfirmPassword.Text) = 0 Then
         lblConfirmPasswordMissing.Visible = True
         Ok = False
      Else
         lblConfirmPasswordMissing.Visible = False
      End If
   End If
   cmdOK.Enabled = Ok
End Sub

Private Sub ImageCombo1_Click()

   Client.CultureLanguage = ImageCombo1.SelectedItem.Key
   Client.Texts.NewLanguage Client.CultureLanguage
End Sub

Private Sub txtConfirmPassword_Change()

   SetEnabled
End Sub

Private Sub txtConfirmPassword_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtNewPassword_Change()

   SetEnabled
End Sub

Private Sub txtNewPassword_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtPassword_Change()

   SetEnabled
End Sub

Private Sub txtPassword_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtUserName_Change()

   SetEnabled
End Sub

Private Sub txtUserName_GotFocus()

   SelectAllText ActiveControl
End Sub
