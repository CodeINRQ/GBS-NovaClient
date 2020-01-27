VERSION 5.00
Begin VB.Form frmPopup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   1170
   ClientLeft      =   1125
   ClientTop       =   8820
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1170
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuPopup 
      Caption         =   "DictList"
      Index           =   0
      Begin VB.Menu mnuDictList 
         Caption         =   "Lås upp"
         Index           =   10
         Tag             =   "1340101"
      End
      Begin VB.Menu mnuDictList 
         Caption         =   "Spårning..."
         Index           =   20
         Tag             =   "1340102"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Edit"
      Index           =   1
      Begin VB.Menu mnuEdit 
         Caption         =   "Kopiera"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "UserList"
      Index           =   2
      Begin VB.Menu mnuUserList 
         Caption         =   "Lås upp"
         Index           =   20
         Tag             =   "1340101"
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Choise(MenuIndex As Integer, ItemIndex As Integer, Id As Long)
Public Id As Long

Private MenuIndex As Integer

Private Sub Form_Load()

   TranslateForm Me
End Sub

Private Sub mnuDictList_Click(Index As Integer)

   RaiseEvent Choise(MenuIndex, Index, Id)
End Sub

Private Sub mnuEdit_Click(Index As Integer)

   RaiseEvent Choise(MenuIndex, Index, Id)
End Sub

Private Sub mnuPopup_Click(Index As Integer)

   MenuIndex = Index
         
End Sub

Private Sub mnuUserList_Click(Index As Integer)

   RaiseEvent Choise(MenuIndex, Index, Id)
End Sub
