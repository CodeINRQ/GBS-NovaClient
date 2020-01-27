VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucOrgTree 
   BackStyle       =   0  'Transparent
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ScaleHeight     =   6540
   ScaleWidth      =   2205
   Begin MSComctlLib.TreeView tvOrgTree 
      Height          =   6495
      HelpContextID   =   1390000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   11456
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   265
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   -120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":00FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":01F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":02EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":03E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":04E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":05DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrgTree.ctx":06D6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ucOrgTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event NewSelect(OrgId As Long, Txt As String)

Dim Level1Id As Long
Dim LastNodeClicked As Node
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
   
Private Sub UserControl_Initialize()

   tvOrgTree.ImageList = ImageList
   tvOrgTree.LabelEdit = tvwManual
   tvOrgTree.LineStyle = tvwRootLines
   tvOrgTree.Style = 7
End Sub

Private Sub UserControl_Resize()

   tvOrgTree.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
Public Property Let Enabled(Value As Boolean)

   tvOrgTree.Enabled = Value
End Property
Public Property Get Enabled() As Boolean

   Enabled = tvOrgTree.Enabled
End Property

Private Sub tvOrgTree_NodeClick(ByVal Node As MSComctlLib.Node)

   Dim SelectedKey As String

   If Len(Node.Tag) > 0 Then
      If Node.Selected Then
         If Not LastNodeClicked Is Nothing Then
            LastNodeClicked.Bold = False
         End If
         Set LastNodeClicked = Node
         Node.Bold = True
         SelectedKey = mId$(Node.Key, 2)
         'tvOrgTree.ToolTipText = SelectedKey
         RaiseEvent NewSelect(CLng(SelectedKey), Node.Text)
      End If
   Else
      If Not LastNodeClicked Is Nothing Then
         LastNodeClicked.Selected = True
      End If
   End If
End Sub

Public Sub AddNode(Id As Long, Parent As Long, Txt As String, Color As Integer, Enabled As Boolean)

   Dim Nod As Node
   
   If Parent = 0 Then
      Set Nod = tvOrgTree.Nodes.Add(, , "_" & CStr(Id), Txt, Color, Color + 1)
   Else
      Set Nod = tvOrgTree.Nodes.Add("_" & CStr(Parent), tvwChild, "_" & CStr(Id), Txt, Color, Color + 1)
   End If
   If Enabled Then
      Nod.Tag = "1"
   Else
      Nod.Tag = ""
   End If
End Sub

Public Sub PickOrgId(OrgId As Long)

   Dim Nod As Node
   
   For Each Nod In tvOrgTree.Nodes
      If CLng(mId$(Nod.Key, 2)) = OrgId Then
         Nod.Selected = True
         tvOrgTree_NodeClick Nod
         Exit For
      End If
   Next Nod
End Sub
Public Sub CloaseAll()

   Dim Nod As Node
   
   For Each Nod In tvOrgTree.Nodes
      If Nod.Selected Then
         Nod.Selected = False
      End If
   Next Nod
End Sub
Public Sub Clear()

   tvOrgTree.Nodes.Clear
   Level1Id = 0
   Set LastNodeClicked = Nothing
End Sub

