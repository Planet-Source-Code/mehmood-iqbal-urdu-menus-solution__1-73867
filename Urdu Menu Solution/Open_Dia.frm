VERSION 5.00
Begin VB.Form Open_Dia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   Icon            =   "Open_Dia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   2760
      Picture         =   "Open_Dia.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   5400
      Picture         =   "Open_Dia.frx":0BE5
      Top             =   240
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   480
      Picture         =   "Open_Dia.frx":129E
      Top             =   360
      Width           =   4680
   End
End
Attribute VB_Name = "Open_Dia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

