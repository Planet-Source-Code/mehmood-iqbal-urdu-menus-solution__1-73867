VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Urdu Menu's Sloution"
   ClientHeight    =   5820
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8550
   Icon            =   "Main_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Labels Events :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   4095
      Begin VB.Frame Frame6 
         Caption         =   "Mouse Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   1815
         Begin VB.Label Label4 
            Caption         =   "Second Label Control"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Click :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1935
         Begin MSForms.Label Label3 
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1335
            Caption         =   "First Label Control "
            Size            =   "2355;450"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Command Button Events :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      Begin VB.Frame Frame4 
         Caption         =   "Mouse Over  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton Command2 
            Caption         =   "<Mouse Over >"
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Click :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton Command1 
            Caption         =   "Right Click >>"
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Image Image11 
      Height          =   420
      Left            =   4920
      Picture         =   "Main_Form.frx":0A02
      Top             =   2880
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image12 
      Height          =   420
      Left            =   4920
      Picture         =   "Main_Form.frx":0C06
      Top             =   3240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image13 
      Height          =   420
      Left            =   4920
      Picture         =   "Main_Form.frx":0E27
      Top             =   3600
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image14 
      Height          =   420
      Left            =   4920
      Picture         =   "Main_Form.frx":1046
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image15 
      Height          =   420
      Left            =   4920
      Picture         =   "Main_Form.frx":1254
      Top             =   4320
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image10 
      Height          =   420
      Left            =   3480
      Picture         =   "Main_Form.frx":146F
      Top             =   4680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image9 
      Height          =   420
      Left            =   3480
      Picture         =   "Main_Form.frx":16B4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image8 
      Height          =   420
      Left            =   3480
      Picture         =   "Main_Form.frx":18F2
      Top             =   3720
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   3480
      Picture         =   "Main_Form.frx":1B30
      Top             =   3240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image4 
      Height          =   420
      Left            =   2040
      Picture         =   "Main_Form.frx":1D6F
      Top             =   3840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   2040
      Picture         =   "Main_Form.frx":1F71
      Top             =   3480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2040
      Picture         =   "Main_Form.frx":2190
      Top             =   3120
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   2040
      Picture         =   "Main_Form.frx":240D
      Top             =   2760
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSForms.Label Label1 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Size            =   "8916;1085"
      Picture         =   "Main_Form.frx":262E
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   8775
   End
   Begin VB.Image Image6 
      Height          =   420
      Left            =   3480
      Picture         =   "Main_Form.frx":2D1C
      Top             =   2760
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image5 
      Height          =   420
      Left            =   2040
      Picture         =   "Main_Form.frx":2F3B
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Menu File_Menu 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu Right_Menu 
         Caption         =   "&Options"
         Visible         =   0   'False
         Begin VB.Menu Item11 
            Caption         =   "&Item11"
         End
         Begin VB.Menu Item22 
            Caption         =   "&Item22"
         End
         Begin VB.Menu Item33 
            Caption         =   "&Item33"
         End
         Begin VB.Menu Item44 
            Caption         =   "&Item44"
         End
         Begin VB.Menu Item55 
            Caption         =   "&Item55"
         End
      End
   End
   Begin VB.Menu Message_Menu 
      Caption         =   "&Message"
      Begin VB.Menu M_Item1 
         Caption         =   "&Item1"
      End
      Begin VB.Menu M_Item2 
         Caption         =   "&Item2"
      End
      Begin VB.Menu M_Item3 
         Caption         =   "&Item3"
      End
      Begin VB.Menu M_Item4 
         Caption         =   "&Item4"
      End
      Begin VB.Menu M_Item5 
         Caption         =   "&Item5"
      End
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Urdu Menu Solution is there. This is about how  ''
'' to make & handle with Urdu Menus using VB6.     ''
'' This project will fullfill your requirments to  ''
'' Make Urdu Menus & Pop-Ups.                      ''
'' Keep in mind that it is a Start to make Urdu    ''
'' Menus. Some Enhancements in this Menu project   ''
'' Needed, But at start, i think, it is enough.    ''
''                                                 ''
'' It was not easy to simplify this code for       ''
'' Beginners but I've fully tried to do it.        ''
'' For that reason, more time gained by this       ''
'' Project. But, i hope. Nobody will feel          ''
'' dificulty to understand it.                     ''
''                                                 ''''''''''''''''
'' You can use & modify it in you projects.        ''''''''''''''''
'' I hope, you'll like also this effort. I'll wait for your      ''
'' Comments & Votes on Planet-Source-Code.Com & on my email.     ''
''                                                               ''
'' Regards: Muhammad Mehmood Iqbal   E-Mail: ME_IQ_TM@Yahoo.Com  ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' API Diclarations

'Data Types
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

'Private Function Diclarations
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long

'Private Constants
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

Private Sub Form_Load()

    'Setting Top-Menu Masks
    
    'Menu Where Array 0
    SetMenuBitmap Me, Array(0, 0), Image1.Picture
    SetMenuBitmap Me, Array(0, 1), Image2.Picture
    SetMenuBitmap Me, Array(0, 2), Image3.Picture
    SetMenuBitmap Me, Array(0, 3), Image4.Picture
    SetMenuBitmap Me, Array(0, 4), Image5.Picture
    
    'Menu Where Array 1
    SetMenuBitmap Me, Array(1, 0), Image6.Picture
    SetMenuBitmap Me, Array(1, 1), Image7.Picture
    SetMenuBitmap Me, Array(1, 2), Image8.Picture
    SetMenuBitmap Me, Array(1, 3), Image9.Picture
    SetMenuBitmap Me, Array(1, 4), Image10.Picture
    
    
End Sub

Private Sub Form_Resize()

'Moving Label2 at buttom of the Main_Form
Label2.Move 0, ScaleHeight - Label2.Height, ScaleWidth
    
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'Showing Same Menu like Form_Main at Label1
Form_MouseDown Button, Shift, X, Y
    
End Sub
Private Sub Command1_Click()

'Calling Private Function For Menu
Call Controls_Menu

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Calling Private Function For Menu
Call Controls_Menu

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'If Right Button from Mouse Pressed at Form_Main Then Show Menu
If Button = vbRightButton Then
    
    'Calling Private Function For Menu
    Call Controls_Menu
    
End If
    
End Sub

Private Sub Label3_Click()

'Calling Private Function For Menu
Call Controls_Menu

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Calling Private Function For Menu
Call Controls_Menu

End Sub

Private Sub M_Item1_Click()

   'Showing Urdu Message Form U_Msg Form
   U_Msg.Show
   
End Sub

Private Sub M_Item2_Click()
 
    'Showing Urdu Message Form U_Msg Form
    U_Msg.Show
    
End Sub

Private Sub M_Item3_Click()

    'Showing Urdu Message Form U_Msg Form
    U_Msg.Show
    
End Sub

Private Sub M_Item4_Click()

    'Showing Urdu Message Form U_Msg Form
    U_Msg.Show
    
End Sub

Private Sub M_Item5_Click()

    'Showing Urdu Message Form U_Msg Form
    U_Msg.Show
    
End Sub

Private Sub Item11_Click()

    'Changing Label2 Caption When Right Click Menu Item, Clicked
    Label2.Caption = "Urdu Menu 1"
    
End Sub


Private Sub Item22_Click()

    'Changing Label2 Caption When Right Click Menu Item, Clicked
    Label2.Caption = "Urdu Menu 2"
    
End Sub


Private Sub Item33_Click()

    'Changing Label2 Caption When Right Click Menu Item, Clicked
    Label2.Caption = "Urdu Menu 3"
    
End Sub


Private Sub Item44_Click()

    'Changing Label2 Caption When Right Click Menu Item, Clicked
    Label2.Caption = "Urdu Menu 4"
    
End Sub


Private Sub Item55_Click()

    'Changing Label2 Caption When Right Click Menu Item, Clicked
    Label2.Caption = "Urdu Menu 5"
    
End Sub

Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)

'Private Function To Put A Bitmap In A Menu Item (As Mask).
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

'Getting the Menu Handle
menu_handle = GetMenu(frm.hwnd)

    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

'Initialize The Menu Properties
With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
End With

'Assign the picture.
SetMenuItemInfo menu_handle, item_numbers(UBound(item_numbers)), True, menu_info
        
End Sub

Private Sub Controls_Menu()

'Setting Top-Menu Masks Form Image Controls
Right_Menu.Visible = True
        
        SetMenuBitmap Me, Array(0, 5, 0), Image11.Picture
        SetMenuBitmap Me, Array(0, 5, 1), Image12.Picture
        SetMenuBitmap Me, Array(0, 5, 2), Image13.Picture
        SetMenuBitmap Me, Array(0, 5, 3), Image14.Picture
        SetMenuBitmap Me, Array(0, 5, 4), Image15.Picture
        
'Displaying Pop-Up Menu.
PopupMenu Right_Menu

'Hiding Pop-Up Menu
Right_Menu.Visible = False

End Sub

Private Sub New_Click()

'New Clicked From Top Menu Then
New_Dia.Show

End Sub

Private Sub Open_Click()

'Open Clicked From Top Menu Then
Open_Dia.Show

End Sub

Private Sub Print_Click()

'Print Clicked From Top Menu Then
Print_Dia.Show

End Sub

Private Sub Save_Click()

'Save Clicked From Top Menu Then
Save_Dia.Show

End Sub

Private Sub Exit_Click()

'Show Close Confirmation Dialog
Close_Con.Show

End Sub

'End Of The Whole Project
