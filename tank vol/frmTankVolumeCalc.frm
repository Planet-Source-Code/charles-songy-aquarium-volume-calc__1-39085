VERSION 5.00
Begin VB.Form frmTankVolumeCalc 
   Caption         =   "Aquarium Volume Calculator"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   Picture         =   "frmTankVolumeCalc.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picVolume 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   1200
      ScaleHeight     =   1515
      ScaleWidth      =   5595
      TabIndex        =   7
      Top             =   2760
      Width           =   5655
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtDepth 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtLength 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblDepth 
      BackColor       =   &H00000000&
      Caption         =   "Depth in inches."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Caption         =   "Width in inches."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLength 
      BackColor       =   &H00000000&
      Caption         =   "Length in inches."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmTankVolumeCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()

Dim x As Single
Dim y As Single
Dim z As Single
x = Val(txtLength.Text)
y = Val(txtWidth.Text)
z = Val(txtDepth.Text)
Call calcvol(x, y, z)

End Sub
Private Sub calcvol(l, w, d As Single)

Dim vol As Single
picVolume.Cls
vol = l * w * d
picVolume.Print "The tank is"; vol; "cubic inches."
Call calcgal(vol)

End Sub
Private Sub calcgal(x As Single)
Dim vol As Single
vol = x * 0.00433
picVolume.Print "The tank holds"; vol; "gallons."

End Sub

Private Sub lblNote_Click()
MsgBox ("Length is the measurement from the front to the back of the tank, width is the measurement from right to left and the depth is the height of the water. Be sure to take into consideration how much substrate you will have.")
End Sub

Private Sub txtLength_gotfocus()
txtLength.Text = ""
txtWidth.Text = ""
txtDepth.Text = ""
End Sub
