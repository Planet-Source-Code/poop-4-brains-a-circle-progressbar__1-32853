VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCircle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circle Progress Bar"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Colors"
      Height          =   2535
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   2535
      Begin MSComDlg.CommonDialog cmn 
         Left            =   480
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   18
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox bckColor 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.PictureBox BColor 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox TColor 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Background  Color"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Non-Value Color"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Value Color"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Values"
      Height          =   2055
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   2295
      Begin MSComctlLib.Slider sldMax 
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Min             =   5
         Max             =   50
         SelStart        =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldValue 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
      End
      Begin MSComctlLib.Slider sldRad 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Min             =   5
         Max             =   50
         SelStart        =   25
         Value           =   25
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress Bar Radius"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.14159

Private Sub cmdColor_Click(Index As Integer)
Select Case Index
Case 0: ChangeColor TColor
Case 1: ChangeColor BColor
Case 2: ChangeColor bckColor
End Select
End Sub

Function ChangeColor(pic As PictureBox)
cmn.DialogTitle = "Change Color"
cmn.Color = pic.BackColor
cmn.ShowColor

pic.BackColor = cmn.Color
End Function

Private Sub cmdDraw_Click()
Board.Cls
Board.BackColor = bckColor.BackColor

Dim I As Long, per, xs, ys, cx, cy

per = sldValue.Value / sldMax.Value * 100
per = per / 100
per = 360 * per

cx = Board.ScaleWidth \ 2
cy = Board.ScaleHeight \ 2

Board.DrawWidth = 2

For I = 0 To 360
xs = Cos(I / 180 * PI) * sldRad.Value
ys = Sin(I / 180 * PI) * sldRad.Value
Board.Line (cx, cy)-(cx + xs, cy + ys), BColor.BackColor
DoEvents
Next I

For I = 0 To per
xs = Cos(I / 180 * PI) * sldRad.Value
ys = Sin(I / 180 * PI) * sldRad.Value
Board.Line (cx, cy)-(cx + xs, cy + ys), TColor.BackColor
DoEvents
Next I
End Sub

Private Sub sldMax_Click()
sldValue.Max = sldMax.Value
End Sub
