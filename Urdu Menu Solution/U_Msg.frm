VERSION 5.00
Begin VB.Form U_Msg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   Icon            =   "U_Msg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   2160
      Picture         =   "U_Msg.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4440
      Picture         =   "U_Msg.frx":0BE5
      Top             =   360
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   360
      Picture         =   "U_Msg.frx":1341
      Top             =   360
      Width           =   3825
   End
End
Attribute VB_Name = "U_Msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

