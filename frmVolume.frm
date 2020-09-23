VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVolume 
   BorderStyle     =   0  'None
   Caption         =   "Volume"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab VolumeTab 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Cube"
      TabPicture(0)   =   "frmVolume.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCube"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCubeL"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCubeAns"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCubeReset"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Cylinder"
      TabPicture(1)   =   "frmVolume.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdCylinder"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtCylinderRadius"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtCylinderAns"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCylinderHeight"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdCylinderReset"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Sphere"
      TabPicture(2)   =   "frmVolume.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSphere"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtSphereRadius"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtSphereAns"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdSphereReset"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "T Prism"
      TabPicture(3)   =   "frmVolume.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label12"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label13"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label14"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdPrismReset"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtPrismAns"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtPrismB"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdPrism"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtPrismH"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Cone"
      TabPicture(4)   =   "frmVolume.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label26"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Cuboid"
      TabPicture(5)   =   "frmVolume.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label5"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label4"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label3"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label10"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "txtCuboidH"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "txtCuboidB"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "cmdCuboid"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "txtCuboidL"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "txtCuboidAns"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "cmdCuboidReset"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).ControlCount=   10
      Begin VB.TextBox txtPrismH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         TabIndex        =   35
         Text            =   "0"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrism 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -71760
         TabIndex        =   34
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtPrismB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         TabIndex        =   33
         Text            =   "0"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtPrismAns 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrismReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73320
         TabIndex        =   31
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdCuboidReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73440
         TabIndex        =   30
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdSphereReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73440
         TabIndex        =   29
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdCylinderReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73440
         TabIndex        =   28
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdCubeReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1800
         TabIndex        =   27
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtCylinderHeight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   25
         Text            =   "0"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtCuboidAns 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtSphereAns 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtCylinderAns 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtCubeAns 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtSphereRadius 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72360
         TabIndex        =   10
         Text            =   "0"
         Top             =   1500
         Width           =   1935
      End
      Begin VB.CommandButton cmdSphere 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -71880
         TabIndex        =   9
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtCylinderRadius 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   8
         Text            =   "0"
         Top             =   1380
         Width           =   1935
      End
      Begin VB.CommandButton cmdCylinder 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -71880
         TabIndex        =   7
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtCuboidL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   6
         Text            =   "0"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdCuboid 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -71880
         TabIndex        =   5
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtCuboidB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   4
         Text            =   "0"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtCuboidH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   3
         Text            =   "0"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtCubeL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdCube 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Avilable in next version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72840
         TabIndex        =   39
         Top             =   2160
         Width           =   2265
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73080
         TabIndex        =   38
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Base :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72960
         TabIndex        =   37
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Answer :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73200
         TabIndex        =   36
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label11 
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   26
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Answer :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73320
         TabIndex        =   24
         Top             =   2760
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Answer :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73440
         TabIndex        =   22
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Answer :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73320
         TabIndex        =   20
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Answer :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   18
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Radius :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   16
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Radius :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73560
         TabIndex        =   15
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Length :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73200
         TabIndex        =   14
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Breadth :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73320
         TabIndex        =   13
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73200
         TabIndex        =   12
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Length :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1920
         TabIndex        =   11
         Top             =   1440
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCube_Click()
On Error GoTo handler
txtCubeAns.Text = txtCubeL ^ 3
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdCubeReset_Click()
txtCubeL.Text = 0
txtCubeAns.Text = 0
End Sub

Private Sub cmdCuboid_Click()
On Error GoTo handler
txtCuboidAns.Text = txtCuboidL.Text * txtCuboidB.Text * txtCuboidH.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdCuboidReset_Click()
txtCuboidL.Text = 0
txtCuboidB.Text = 0
txtCuboidH.Text = 0
txtCuboidAns.Text = 0
End Sub

Private Sub cmdCylinder_Click()
On Error GoTo handler
txtCylinderAns.Text = 3.1416 * (txtCylinderRadius.Text ^ 2) * txtCylinderHeight.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdCylinderReset_Click()
txtCylinderRadius.Text = 0
txtCylinderHeight.Text = 0
txtCylinderAns.Text = 0
End Sub

Private Sub cmdPrism_Click()
On Error GoTo handler
txtPrismAns.Text = (1 / 2 * txtPrismB.Text * txtPrismH.Text) * txtPrismH.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdPrismReset_Click()
txtPrismAns.Text = 0
txtPrismB.Text = 0
txtPrismH.Text = 0
End Sub

Private Sub cmdSphere_Click()
On Error GoTo handler
txtSphereAns.Text = (4 * 3.1416 * (txtSphereRadius.Text ^ 3)) / 4
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdSphereReset_Click()
txtSphereRadius.Text = 0
txtSphereAns.Text = 0
End Sub

Private Sub Form_Load()
Me.Icon = Calculator.Icon
End Sub

Private Sub txtCubeL_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCubeL.Text
End Sub

Private Sub txtCuboidL_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCuboidL.Text
End Sub

Private Sub txtCuboidB_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCuboidB.Text
End Sub

Private Sub txtCuboidH_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCuboidH.Text
End Sub

Private Sub txtCylinderHeight_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCylinderHeight.Text
End Sub

Private Sub txtCylinderRadius_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCylinderRadius.Text
End Sub

Private Sub txtPrismAns_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtPrismAns.Text
End Sub

Private Sub txtPrismB_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtPrismB.Text
End Sub

Private Sub txtPrismH_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtPrismH.Text
End Sub

Private Sub txtSphereRadius_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtSphereRadius.Text
End Sub
