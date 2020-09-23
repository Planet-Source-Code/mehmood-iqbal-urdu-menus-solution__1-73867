VERSION 5.00
Begin VB.Form New_Dia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   Icon            =   "New_Dia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   2040
      Picture         =   "New_Dia.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4200
      Picture         =   "New_Dia.frx":0BE5
      Top             =   360
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   -480
      Picture         =   "New_Dia.frx":12AB
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "New_Dia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
