VERSION 5.00
Begin VB.Form Save_Dia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "Save_Dia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   2040
      Picture         =   "Save_Dia.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4560
      Picture         =   "Save_Dia.frx":0BE5
      Top             =   240
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   120
      Picture         =   "Save_Dia.frx":127B
      Top             =   360
      Width           =   4680
   End
End
Attribute VB_Name = "Save_Dia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
