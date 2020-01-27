VERSION 5.00
Begin VB.Form frmEditDictType 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diktattyp"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6030
   HelpContextID   =   1040000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1430100"
   Begin VB.TextBox txtDictTypeId 
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtDictTypeText 
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "1430102"
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Tag             =   "1430104"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Spara"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Tag             =   "1430103"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblDictTypeIdTitle 
      Caption         =   "&Id:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "1430101"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDictTypeTextTitle 
      Caption         =   "&Text:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Tag             =   "1040101"
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditDictType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public DictTypeToEdit As clsDictType
Public Event SaveClicked()

Private Dirty As Boolean

Private Sub cmdCancel_Click()

   Unload Me
End Sub

Private Sub cmdSave_Click()

   DictTypeToEdit.DictTypeId = CInt(txtDictTypeId.Text)
   DictTypeToEdit.DictTypeText = txtDictTypeText.Text
   Client.DictTypeMgr.SaveDictType DictTypeToEdit
   
   RaiseEvent SaveClicked
   Unload Me
End Sub

Private Sub Form_Activate()

   If DictTypeToEdit.DictTypeId < 0 Then
      txtDictTypeId.Text = ""
      txtDictTypeId.Enabled = True
   Else
      txtDictTypeId.Text = CStr(DictTypeToEdit.DictTypeId)
      txtDictTypeId.Enabled = False
   End If
   txtDictTypeText.Text = DictTypeToEdit.DictTypeText
End Sub

Private Sub SetEnabled()

   cmdSave.Enabled = Dirty And Len(txtDictTypeId.Text) > 0
End Sub

Private Sub Form_Load()

   CenterAndTranslateForm Me, frmMain

   SetEnabled
End Sub

Private Sub txtDictTypeId_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtDictTypeId_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtDictTypeId_KeyPress(KeyAscii As Integer)

   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii < 32) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDictTypeText_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub Label2_Click()

End Sub

Private Sub txtDictTypeText_GotFocus()

   SelectAllText ActiveControl
End Sub
