VERSION 5.00
Begin VB.UserControl ucVUmeter 
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ScaleHeight     =   90
   ScaleWidth      =   3105
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
   Begin VB.Shape shpBacground 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   375
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

Private Sub UserControl_Initialize()

   FullWidth = UserControl.Width
End Sub
Public Property Let Value(Value As Integer)

   mValue = Value
   UserControl.Width = CInt(CSng(FullWidth) * (mValue / 100))
End Property
Public Property Get Value() As Integer

   Value = mValue
End Property

