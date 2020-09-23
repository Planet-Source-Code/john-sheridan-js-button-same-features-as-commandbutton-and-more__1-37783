VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "John's Button OCX"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Change Fade"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Backcolor"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin Project1.JSbutton JSbutton1 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
   End
   Begin VB.Label Label2 
      Caption         =   "Note: You can change the BackColor and the fade amount in code."
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
JSbutton1.BackColor = vbBlue
End Sub

Private Sub Command2_Click()
JSbutton1.fadeAmount = 100
End Sub


Private Sub JSbutton1_clickMe()
MsgBox "Hey! Get off my button!", vbCritical, "grrr..."
End Sub

