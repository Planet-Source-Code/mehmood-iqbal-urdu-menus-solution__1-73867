VERSION 5.00
Begin VB.Form Close_Con 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "Close_Con.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   3360
      Picture         =   "Close_Con.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   1680
      Picture         =   "Close_Con.frx":0B7A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   360
      Picture         =   "Close_Con.frx":0D20
      Top             =   360
      Width           =   4680
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   5160
      Picture         =   "Close_Con.frx":1461
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Close_Con"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
Unload Main_Form
End Sub
