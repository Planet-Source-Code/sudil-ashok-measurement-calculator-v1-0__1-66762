VERSION 5.00
Begin VB.MDIForm MDIFormMain 
   BackColor       =   &H8000000F&
   Caption         =   "RS Measurement Calculator"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   Icon            =   "MDIFormMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9885
      TabIndex        =   3
      Top             =   4950
      Width           =   9945
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Copyright © Regider Software|www.Regider.co.nr|regider@gmail.com|Developer Sudil && Ashok"
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
         Left            =   270
         TabIndex        =   4
         Top             =   0
         Width           =   9330
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   750
         Left            =   1200
         Picture         =   "MDIFormMain.frx":000C
         ScaleHeight     =   526.75
         ScaleMode       =   0  'User
         ScaleWidth      =   526.75
         TabIndex        =   2
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RS Measurement Calculator"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   6195
      End
   End
End
Attribute VB_Name = "MDIFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Calculator.Show
Me.Icon = Calculator.Icon
End Sub

Private Sub MDIForm_Resize()
If WindowState = 2 Then WindowState = 0
Me.Height = 7000
Me.Width = 10065
End Sub
