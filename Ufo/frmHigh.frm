VERSION 5.00
Begin VB.Form frmHigh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmHigh.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl10 
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      Height          =   255
      Left            =   60
      TabIndex        =   29
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lbl09 
      BackStyle       =   0  'Transparent
      Caption         =   "9."
      Height          =   255
      Left            =   60
      TabIndex        =   28
      Top             =   2460
      Width           =   255
   End
   Begin VB.Label lbl08 
      BackStyle       =   0  'Transparent
      Caption         =   "8."
      Height          =   255
      Left            =   60
      TabIndex        =   27
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lbl07 
      BackStyle       =   0  'Transparent
      Caption         =   "7."
      Height          =   255
      Left            =   60
      TabIndex        =   26
      Top             =   1860
      Width           =   255
   End
   Begin VB.Label lbl06 
      BackStyle       =   0  'Transparent
      Caption         =   "6."
      Height          =   255
      Left            =   60
      TabIndex        =   25
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lbl05 
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   1260
      Width           =   255
   End
   Begin VB.Label lbl04 
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lbl03 
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      Height          =   255
      Left            =   60
      TabIndex        =   22
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lbl02 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lbl01 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   60
      Width           =   255
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   2700
      TabIndex        =   19
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   18
      Top             =   2760
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   2700
      TabIndex        =   17
      Top             =   2460
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   16
      Top             =   2460
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   2700
      TabIndex        =   15
      Top             =   2160
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   2700
      TabIndex        =   13
      Top             =   1860
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   12
      Top             =   1860
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   11
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   2700
      TabIndex        =   9
      Top             =   1260
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   1260
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   2700
      TabIndex        =   7
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   2700
      TabIndex        =   5
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   660
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   2700
      TabIndex        =   3
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2235
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   2700
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   60
      Width           =   2235
   End
End
Attribute VB_Name = "frmHigh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    OnTop hWnd, True
End Sub
