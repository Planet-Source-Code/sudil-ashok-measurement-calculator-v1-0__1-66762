VERSION 5.00
Begin VB.Form Calculator 
   BorderStyle     =   0  'None
   Caption         =   "RS Measurement Calculator"
   ClientHeight    =   5130
   ClientLeft      =   105
   ClientTop       =   -240
   ClientWidth     =   3000
   Icon            =   "Calculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCalculation 
      Caption         =   "Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdVolume 
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdArea 
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
frmAbout.Show vbModal, MDIFormMain
End Sub

Private Sub cmdArea_Click()
frmVolume.Hide
frmCalculation.Hide
frmArea.Left = 3500
frmArea.Top = 250
frmArea.Show
End Sub

Private Sub cmdCalculation_Click()
frmArea.Hide
frmVolume.Hide
frmCalculation.Left = 3500
frmCalculation.Top = 250
frmCalculation.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdVolume_Click()
frmArea.Hide
frmCalculation.Hide
frmVolume.Left = 3500
frmVolume.Top = 250
frmVolume.Show
End Sub

Sub Form_Unload(cancel As Integer)
End
End Sub
