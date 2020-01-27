VERSION 5.00
Begin VB.UserControl ucCloseChoice 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ScaleHeight     =   1710
   ScaleWidth      =   2625
   Begin VB.Frame fraChoice 
      Caption         =   "Spara som"
      Height          =   1695
      HelpContextID   =   1060000
      Left            =   0
      TabIndex        =   7
      Tag             =   "1060101"
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdClose 
         Caption         =   "Stäng"
         Height          =   275
         Left            =   120
         TabIndex        =   0
         Tag             =   "1060102"
         Top             =   1350
         Width           =   2295
      End
      Begin VB.OptionButton optChoice 
         BackColor       =   &H008080FF&
         Height          =   315
         Index           =   0
         Left            =   240
         Picture         =   "CloseChoice.ctx":0000
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optChoice 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   1
         Left            =   240
         Picture         =   "CloseChoice.ctx":0502
         TabIndex        =   2
         Top             =   615
         Width           =   255
      End
      Begin VB.OptionButton optChoice 
         BackColor       =   &H0080FF80&
         Height          =   315
         Index           =   2
         Left            =   240
         Picture         =   "CloseChoice.ctx":0A04
         TabIndex        =   3
         Top             =   975
         Width           =   255
      End
      Begin VB.Label lblChoiseMissing 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape shpBackground 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   465
      End
      Begin VB.Shape shpBackground 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   465
      End
      Begin VB.Shape shpBackground 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   315
         UseMnemonic     =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   675
         UseMnemonic     =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   1035
         UseMnemonic     =   0   'False
         Width           =   1935
      End
   End
End
Attribute VB_Name = "ucCloseChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event NewSelect(Index As Integer)
Public Event CloseClicked()

Private mChoiceValue As Integer
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Property Let ChoiceText(Index As Integer, Text As String)

   If Index >= 0 And Index <= 2 Then
      lblText(Index).Caption = Text
      If Len(Text) > 0 Then
         optChoice(Index).Enabled = True
         optChoice(Index).Value = True
      Else
         optChoice(Index).Enabled = False
         optChoice(Index).Value = False
      End If
   End If
End Property
Public Property Get ChoiceText(Index As Integer) As String

   If Index >= 0 And Index <= 2 Then
      ChoiceText = lblText(Index).Caption
   End If
End Property
Public Property Let ChoiceValue(Index As Integer)

   If Index >= 0 And Index <= 2 Then
      mChoiceValue = Index
      lblChoiseMissing.Visible = False
      cmdClose.Enabled = True
      optChoice(Index).Value = True
   Else
      mChoiceValue = -1
      optChoice(0).Value = False
      optChoice(1).Value = False
      optChoice(2).Value = False
      lblChoiseMissing.Visible = True
   End If
End Property
Public Property Get ChoiceValue() As Integer

   ChoiceValue = mChoiceValue
End Property
Public Property Let ChoiceTip(Index As Integer, Text As String)

   If Index >= 0 And Index <= 2 Then
      lblText(Index).ToolTipText = Text
      optChoice(Index).ToolTipText = Text
   End If
End Property
Private Sub cmdClose_Click()

   Trc "ucClose Event CloseClicked", ""
   RaiseEvent CloseClicked
End Sub

Private Sub lblText_Click(Index As Integer)

   optChoice(Index).Value = True
End Sub

Private Sub optChoice_Click(Index As Integer)

   If optChoice(Index).Value = True Then
      mChoiceValue = Index
      lblChoiseMissing.Visible = False
      cmdClose.Enabled = True
      Trc "ucClose Event NewSelect", Format$(Index)
      RaiseEvent NewSelect(Index)
   End If
End Sub

Private Sub UserControl_Initialize()

   Trc "ucClose Initialize", ""
   mChoiceValue = -1
   cmdClose.Enabled = True
End Sub

Private Sub UserControl_Terminate()

   Trc "ucClose Terminate", ""
End Sub
