VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmArea 
   BorderStyle     =   0  'None
   Caption         =   "Area"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab AreaTab 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   12
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Square"
      TabPicture(0)   =   "frmArea.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSquareReset"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSquareAns"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSquare"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSquareS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Rectangle"
      TabPicture(1)   =   "frmArea.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRectReset"
      Tab(1).Control(1)=   "txtRectAns"
      Tab(1).Control(2)=   "txtRectB"
      Tab(1).Control(3)=   "cmdRect"
      Tab(1).Control(4)=   "txtRectL"
      Tab(1).Control(5)=   "Label8"
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(7)=   "Label3"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Triangle"
      TabPicture(2)   =   "frmArea.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTriangleAns1"
      Tab(2).Control(1)=   "txtTriangleS3"
      Tab(2).Control(2)=   "txtTriangleS2"
      Tab(2).Control(3)=   "txtTriangleS1"
      Tab(2).Control(4)=   "cmdTriangleReset"
      Tab(2).Control(5)=   "txtTriangleAns"
      Tab(2).Control(6)=   "txtTriangleB"
      Tab(2).Control(7)=   "cmdTriangle"
      Tab(2).Control(8)=   "txtTriangleA"
      Tab(2).Control(9)=   "Label17"
      Tab(2).Control(10)=   "Label16"
      Tab(2).Control(11)=   "Label15"
      Tab(2).Control(12)=   "Label14"
      Tab(2).Control(13)=   "Label6"
      Tab(2).Control(14)=   "Label5"
      Tab(2).Control(15)=   "Label1"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Circle"
      TabPicture(3)   =   "frmArea.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdCircleReset"
      Tab(3).Control(1)=   "txtCircleAns"
      Tab(3).Control(2)=   "txtCircleRadius"
      Tab(3).Control(3)=   "cmdCircle"
      Tab(3).Control(4)=   "Label7"
      Tab(3).Control(5)=   "Label2"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Parallelogram"
      TabPicture(4)   =   "frmArea.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtParallelH"
      Tab(4).Control(1)=   "cmdParallel"
      Tab(4).Control(2)=   "txtParallelB"
      Tab(4).Control(3)=   "txtParallelAns"
      Tab(4).Control(4)=   "cmdParallelReset"
      Tab(4).Control(5)=   "Label13"
      Tab(4).Control(6)=   "Label12"
      Tab(4).Control(7)=   "Label11"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Cube"
      TabPicture(5)   =   "frmArea.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label26"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Cylinder"
      TabPicture(6)   =   "frmArea.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label27"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Sphere"
      TabPicture(7)   =   "frmArea.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label28"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Hemi-Sphere"
      TabPicture(8)   =   "frmArea.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label29"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Cone"
      TabPicture(9)   =   "frmArea.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label30"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Trapezium"
      TabPicture(10)  =   "frmArea.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "cmdTrap"
      Tab(10).Control(1)=   "cmdTrapReset"
      Tab(10).Control(2)=   "txtTrapL"
      Tab(10).Control(3)=   "txtTrapB"
      Tab(10).Control(4)=   "txtTrapH"
      Tab(10).Control(5)=   "txtTrapAns"
      Tab(10).Control(6)=   "Label21"
      Tab(10).Control(7)=   "Label20"
      Tab(10).Control(8)=   "Label19"
      Tab(10).Control(9)=   "Label18"
      Tab(10).ControlCount=   10
      TabCaption(11)  =   "Quadrilateral"
      TabPicture(11)  =   "frmArea.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "txtQuadAns"
      Tab(11).Control(1)=   "txtQuadH2"
      Tab(11).Control(2)=   "txtQuadH1"
      Tab(11).Control(3)=   "txtQuadD"
      Tab(11).Control(4)=   "cmdQuadReset"
      Tab(11).Control(5)=   "cmdQuad"
      Tab(11).Control(6)=   "Label25"
      Tab(11).Control(7)=   "Label24"
      Tab(11).Control(8)=   "Label23"
      Tab(11).Control(9)=   "Label22"
      Tab(11).ControlCount=   10
      Begin VB.TextBox txtQuadAns 
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
         TabIndex        =   60
         Tag             =   "0"
         Text            =   "0"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtQuadH2 
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
         TabIndex        =   59
         Text            =   "0"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtQuadH1 
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
         TabIndex        =   58
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtQuadD 
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
         TabIndex        =   57
         Text            =   "0"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdQuadReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73440
         TabIndex        =   56
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdQuad 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -71880
         TabIndex        =   55
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrap 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -72120
         TabIndex        =   54
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrapReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73680
         TabIndex        =   53
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtTrapL 
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
         Left            =   -72480
         TabIndex        =   45
         Text            =   "0"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtTrapB 
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
         Left            =   -72480
         TabIndex        =   46
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtTrapH 
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
         Left            =   -72480
         TabIndex        =   48
         Text            =   "0"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtTrapAns 
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
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   47
         Tag             =   "0"
         Text            =   "0"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtTriangleAns1 
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
         Left            =   -70800
         Locked          =   -1  'True
         TabIndex        =   18
         Tag             =   "0"
         Text            =   "0"
         Top             =   2940
         Width           =   1575
      End
      Begin VB.TextBox txtTriangleS3 
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
         Left            =   -70800
         TabIndex        =   17
         Text            =   "0"
         Top             =   2340
         Width           =   1575
      End
      Begin VB.TextBox txtTriangleS2 
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
         Left            =   -70800
         TabIndex        =   16
         Text            =   "0"
         Top             =   1860
         Width           =   1575
      End
      Begin VB.TextBox txtTriangleS1 
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
         Left            =   -70800
         TabIndex        =   15
         Text            =   "0"
         Top             =   1380
         Width           =   1575
      End
      Begin VB.TextBox txtParallelH 
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
         Left            =   -72600
         TabIndex        =   34
         Text            =   "0"
         Top             =   1980
         Width           =   1935
      End
      Begin VB.CommandButton cmdParallel 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -72480
         TabIndex        =   36
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox txtParallelB 
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
         Left            =   -72600
         TabIndex        =   33
         Text            =   "0"
         Top             =   1500
         Width           =   1935
      End
      Begin VB.TextBox txtParallelAns 
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   2580
         Width           =   1935
      End
      Begin VB.CommandButton cmdParallelReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -74040
         TabIndex        =   38
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox txtSquareS 
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
         Left            =   2640
         TabIndex        =   1
         Text            =   "0"
         Top             =   1740
         Width           =   1935
      End
      Begin VB.CommandButton cmdSquare 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   3780
         Width           =   1455
      End
      Begin VB.TextBox txtSquareAns 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   2340
         Width           =   1935
      End
      Begin VB.CommandButton cmdSquareReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   3780
         Width           =   1455
      End
      Begin VB.CommandButton cmdCircleReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73800
         TabIndex        =   30
         Top             =   3660
         Width           =   1455
      End
      Begin VB.CommandButton cmdTriangleReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73920
         TabIndex        =   14
         Top             =   3780
         Width           =   1455
      End
      Begin VB.CommandButton cmdRectReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -73800
         TabIndex        =   9
         Top             =   3780
         Width           =   1455
      End
      Begin VB.TextBox txtRectAns 
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtCircleAns 
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
         Top             =   2340
         Width           =   1935
      End
      Begin VB.TextBox txtTriangleAns 
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
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   2460
         Width           =   1575
      End
      Begin VB.TextBox txtTriangleB 
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
         Left            =   -73680
         TabIndex        =   10
         Text            =   "0"
         Top             =   1380
         Width           =   1575
      End
      Begin VB.CommandButton cmdTriangle 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -72360
         TabIndex        =   13
         Top             =   3780
         Width           =   1455
      End
      Begin VB.TextBox txtTriangleA 
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
         Left            =   -73680
         TabIndex        =   11
         Text            =   "0"
         Top             =   1860
         Width           =   1575
      End
      Begin VB.TextBox txtCircleRadius 
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
         TabIndex        =   20
         Text            =   "0"
         Top             =   1740
         Width           =   1935
      End
      Begin VB.CommandButton cmdCircle 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -72240
         TabIndex        =   19
         Top             =   3660
         Width           =   1455
      End
      Begin VB.TextBox txtRectB 
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
         Left            =   -72600
         TabIndex        =   6
         Text            =   "0"
         Top             =   1980
         Width           =   1935
      End
      Begin VB.CommandButton cmdRect 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   -72120
         TabIndex        =   8
         Top             =   3780
         Width           =   1455
      End
      Begin VB.TextBox txtRectL 
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
         Left            =   -72600
         TabIndex        =   5
         Text            =   "0"
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Label Label30 
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
         Left            =   -73080
         TabIndex        =   69
         Top             =   2520
         Width           =   2265
      End
      Begin VB.Label Label29 
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
         Left            =   -73440
         TabIndex        =   68
         Top             =   2520
         Width           =   2265
      End
      Begin VB.Label Label28 
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
         Left            =   -73200
         TabIndex        =   67
         Top             =   2520
         Width           =   2265
      End
      Begin VB.Label Label27 
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
         Left            =   -73200
         TabIndex        =   66
         Top             =   2520
         Width           =   2265
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
         Left            =   -73080
         TabIndex        =   65
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label Label25 
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
         TabIndex        =   64
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Distance1 (H2)"
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
         Left            =   -73920
         TabIndex        =   63
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Distance1 H1)"
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
         Left            =   -73800
         TabIndex        =   62
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Diagonal"
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
         TabIndex        =   61
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
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
         TabIndex        =   52
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Breadth:"
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
         Left            =   -73560
         TabIndex        =   51
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
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
         TabIndex        =   50
         Top             =   2280
         Width           =   750
      End
      Begin VB.Label Label18 
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
         Left            =   -73560
         TabIndex        =   49
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label17 
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
         Left            =   -71880
         TabIndex        =   44
         Top             =   2940
         Width           =   885
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Side3(c):"
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
         Left            =   -71880
         TabIndex        =   43
         Top             =   2340
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Side2(b):"
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
         Left            =   -71880
         TabIndex        =   42
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Side1(a) :"
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
         Left            =   -71880
         TabIndex        =   41
         Top             =   1380
         Width           =   1020
      End
      Begin VB.Label Label13 
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
         Left            =   -73680
         TabIndex        =   40
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label Label12 
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
         Left            =   -73440
         TabIndex        =   39
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label Label11 
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
         Left            =   -73680
         TabIndex        =   37
         Top             =   2580
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Side :"
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
         Left            =   1680
         TabIndex        =   32
         Top             =   1740
         Width           =   615
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
         Left            =   1560
         TabIndex        =   31
         Top             =   2340
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
         Left            =   -73680
         TabIndex        =   29
         Top             =   2580
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
         Left            =   -73440
         TabIndex        =   28
         Top             =   2340
         Width           =   885
      End
      Begin VB.Label Label6 
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   2460
         Width           =   885
      End
      Begin VB.Label Label5 
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
         Left            =   -74520
         TabIndex        =   26
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Altitude :"
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   1860
         Width           =   915
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
         TabIndex        =   24
         Top             =   1740
         Width           =   975
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
         Left            =   -73680
         TabIndex        =   23
         Top             =   1980
         Width           =   945
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
         Left            =   -73560
         TabIndex        =   22
         Top             =   1500
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCircle_Click()
On Error GoTo handler:
txtCircleAns.Text = 3.1416 * (txtCircleRadius ^ 2)
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdCircleReset_Click()
txtCircleRadius.Text = 0
txtCircleAns.Text = 0
End Sub

Private Sub cmdParallel_Click()
On Error GoTo handler
txtParallelAns.Text = txtParallelB.Text * txtParallelH.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdParallelReset_Click()
txtParallelB.Text = 0
txtParallelH.Text = 0
txtParallelAns.Text = 0
End Sub

Private Sub cmdQuad_Click()
On Error GoTo handler
txtQuadAns.Text = 1 / 2 * (txtQuadD.Text * (txtQuadH1.Text + txtQuadH2.Text))
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdQuadReset_Click()
txtQuadD.Text = 0
txtQuadH1.Text = 0
txtQuadH2.Text = 0
txtQuadAns.Text = 0
End Sub

Private Sub cmdRect_Click()
On Error GoTo handler:
txtRectAns.Text = txtRectL.Text * txtRectB.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdRectReset_Click()
txtRectL.Text = 0
txtRectB.Text = 0
txtRectAns.Text = 0
End Sub

Private Sub cmdSquare_Click()
On Error GoTo handler:
txtSquareAns.Text = txtSquareS.Text ^ 2
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdSquareReset_Click()
txtSquareAns.Text = 0
txtSquareS.Text = 0
End Sub

Private Sub cmdTrap_Click()
On Error GoTo handler
txtTrapAns.Text = 1 / 2 * (txtTrapL.Text + txtTrapB.Text) * txtTrapH.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdTrapB_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, cmdTrapB.Text
End Sub

Private Sub cmdTrapH_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, cmdTrapH.Text
End Sub

Private Sub cmdTrapL_Change()

End Sub

Private Sub cmdTrapL_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, cmdTrapL.Text
End Sub

Private Sub cmdTrapReset_Click()
cmdTrapAns.Text = 0
cmdTrapL.Text = 0
cmdTrapB.Text = 0
cmdTrapH.Text = 0
End Sub

Private Sub cmdTriangle_Click()
On Error GoTo handler
txtTriangleAns.Text = 1 / 2 * (txtTriangleB.Text * txtTriangleA.Text)
txtTriangleAns1.Tag = (Val(txtTriangleS1.Text) + Val(txtTriangleS2.Text) + Val(txtTriangleS3.Text)) / 2

'txtTriangleAns1.Text = (Val(txtTriangleAns1.Tag) * ((Val(txtTriangleAns1.Tag) - Val(txtTriangleS1.Text) * (Val(txtTriangleAns1.Tag) - val(txtTriangleS2.Text)) * (Val(txtTriangleAns1.Tag) - val(txtTriangleS3.Text))) ^ 0.5
txtTriangleAns1.Text = (txtTriangleAns1.Tag * (txtTriangleAns1.Tag - txtTriangleS1.Text) * (txtTriangleAns1.Tag - txtTriangleS2.Text) * (txtTriangleAns1.Tag - txtTriangleS3.Text)) ^ 0.5
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdTriangleReset_Click()
txtTriangleB.Text = 0
txtTriangleA.Text = 0
txtTriangleAns.Text = 0
End Sub

Private Sub Command1_Click()
On Error GoTo handler
cmdTrapAns.Text = 1 / 2 * (cmdTrapL.Text + cmdTrapB.Text) * cmdTrapH.Text
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub Form_Load()
Me.Icon = Calculator.Icon
End Sub

Private Sub txtParallelB_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtParallelB.Text
End Sub

Private Sub txtParallelH_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtParallelH.Text
End Sub

Private Sub txtRectB_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtRectB.Text
End Sub

Private Sub txtRectL_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtRectL.Text
End Sub

Private Sub txtCircleRadius_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtCircleRadius.Text
End Sub

Private Sub txtSquareS_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtSquareS.Text
End Sub

Private Sub txtTriangleA_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtTriangleA.Text
End Sub

Private Sub txtTriangleB_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtTriangleB.Text
End Sub

Private Sub txtTriangleS1_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtTriangleS1.Text
End Sub

Private Sub txtTriangleS2_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtTriangleS2.Text
End Sub

Private Sub txtTriangleS3_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtTriangleS3.Text
End Sub
