VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00373436&
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label source 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image menuHover 
      Height          =   1080
      Index           =   2
      Left            =   360
      Picture         =   "frmMenu.frx":0000
      Top             =   5160
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image menuHover 
      Height          =   1080
      Index           =   1
      Left            =   360
      Picture         =   "frmMenu.frx":BD42
      Top             =   3840
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image menuHover 
      Height          =   1080
      Index           =   0
      Left            =   360
      Picture         =   "frmMenu.frx":17A84
      Top             =   2520
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image menu 
      Height          =   1080
      Index           =   2
      Left            =   360
      Picture         =   "frmMenu.frx":237C6
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Image menu 
      Height          =   1080
      Index           =   1
      Left            =   360
      Picture         =   "frmMenu.frx":2F628
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Image menu 
      Height          =   1080
      Index           =   0
      Left            =   360
      Picture         =   "frmMenu.frx":3B48A
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Image logo 
      Height          =   1710
      Left            =   120
      Picture         =   "frmMenu.frx":472EC
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label restoreMenu 
      BackColor       =   &H00373436&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastMenuOption As Integer

Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
logo.left = Screen.Width - logo.Width - 100
End Sub

Private Sub menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If menuHover(Index).Visible = False Then
    menu(Index).Visible = False
    menuHover(Index).Visible = True
    lastMenuOption = Index
End If
End Sub

Private Sub menuHover_Click(Index As Integer)
Dim formToOpen As Form
Select Case Index
    Case 0
        Set formToOpen = frmClient
    Case 2
        Set formToOpen = frmInvoice
End Select

If formToOpen Is Nothing Then Exit Sub

formToOpen.Show , Me
formToOpen.Top = Me.source.Top
formToOpen.left = Me.source.left
End Sub

Private Sub restoreMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If menuHover(lastMenuOption).Visible = True Then
    menu(lastMenuOption).Visible = True
    menuHover(lastMenuOption).Visible = False
End If
End Sub
