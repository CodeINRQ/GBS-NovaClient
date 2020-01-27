VERSION 5.00
Begin VB.UserControl ucVUmeter 
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ScaleHeight     =   150
   ScaleWidth      =   6855
   Begin VB.PictureBox picFrame 
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox picVUmeter 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   6855
         TabIndex        =   1
         Top             =   0
         Width           =   6855
         Begin VB.Shape shpBacground 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   5
            Left            =   2760
            Top             =   0
            Width           =   375
         End
         Begin VB.Shape shpBacground 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   4
            Left            =   2520
            Top             =   0
            Width           =   255
         End
         Begin VB.Shape shpBacground 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0080FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   3
            Left            =   2040
            Top             =   0
            Width           =   495
         End
         Begin VB.Shape shpBacground 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0080FF80&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   2
            Left            =   360
            Top             =   0
            Width           =   1695
         End
         Begin VB.Shape shpBacground 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "ucVUmeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FullWidth  As Integer
Private mValue As Integer

Public Property Let Value(Value As Integer)

   mValue = Value
   SetValue
End Property
Public Property Get Value() As Integer

   Value = mValue
End Property

Private Sub UserControl_Resize()

   SetSize
End Sub
Private Sub SetSize()

   Dim W As Integer
   Dim L As Integer
   Dim I As Integer
   
   FullWidth = UserControl.Width
   picVUmeter.Height = UserControl.Height
   picFrame.Height = UserControl.Height
   picFrame.Width = FullWidth
   For I = 1 To 5
      If I = 2 Then
         W = CInt(FullWidth * 0.6)
      Else
         W = CInt(FullWidth * 0.1)
      End If
      shpBacground(I).Left = L
      shpBacground(I).Width = W + 5  'some extra to avoid gap
      L = L + W
      
      shpBacground(I).Height = UserControl.Height
   Next I
   SetValue
End Sub
Private Sub SetValue()

   picVUmeter.Width = CInt(CSng(FullWidth) * (mValue / 100))
End Sub
