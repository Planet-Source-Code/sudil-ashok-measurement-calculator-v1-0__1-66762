VERSION 5.00
Begin VB.Form frmCalculation 
   BorderStyle     =   0  'None
   Caption         =   "Calculation"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option6 
      Caption         =   "Factorial of"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Power of"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Division"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Multiplication"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Subtraction"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Addition"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtans 
      Alignment       =   2  'Center
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
      Left            =   1440
      TabIndex        =   2
      Text            =   "0"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtsecond 
      Alignment       =   2  'Center
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
      Left            =   1440
      TabIndex        =   1
      Text            =   "0"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtfirst 
      Alignment       =   2  'Center
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
      Left            =   1440
      TabIndex        =   0
      Text            =   "0"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Label labelsign 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1725
   End
End
Attribute VB_Name = "frmCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
On Error GoTo handler
If Option1.Value = True Then
txtans = txtfirst.Text + txtsecond.Text
End If
If Option2.Value = True Then
txtans = txtfirst.Text - txtsecond.Text
End If
If Option3.Value = True Then
txtans = txtfirst.Text * txtsecond.Text
End If
If Option4.Value = True Then
txtans = txtfirst.Text / txtsecond.Text
End If
If Option5.Value = True Then
txtans = txtsecond.Text ^ txtfirst.Text
End If
If Option6.Value = True Then
txtans.Text = 1
For i = 1 To txtfirst.Text
txtans.Text = Val(txtans.Text) * Val(i)
Next i
End If
Exit Sub
handler:
MsgBox "Error: " & Error, vbCritical, "RS Measurement Calculator"
End Sub

Private Sub cmdReset_Click()
txtfirst.Text = 0
txtsecond.Text = 0
txtans.Text = 0
End Sub

Private Sub Option1_Click()
labelsign.Visible = True
txtsecond.Visible = True
labelsign.Caption = "+"
End Sub

Private Sub Option2_Click()
labelsign.Visible = True
txtsecond.Visible = True
labelsign.Caption = "-"
End Sub

Private Sub Option3_Click()
labelsign.Visible = True
txtsecond.Visible = True
labelsign.Caption = "x"
End Sub

Private Sub Option4_Click()
labelsign.Visible = True
txtsecond.Visible = True
labelsign.Caption = "÷"
End Sub

Private Sub Option5_Click()
labelsign.Visible = True
txtsecond.Visible = True
labelsign.Caption = "th power of"
End Sub

Private Sub Option6_Click()
labelsign.Visible = False
txtsecond.Visible = False
End Sub

Private Sub txtfirst_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtfirst.Text
End Sub
Private Sub txtsecond_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtsecond.Text
End Sub

