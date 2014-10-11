VERSION 5.00
Begin VB.Form BG 
   BorderStyle     =   0  'None
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13035
   Icon            =   "bg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   9945
      Left            =   480
      Picture         =   "bg.frx":0CFA
      Top             =   480
      Width           =   15000
   End
End
Attribute VB_Name = "BG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Screen.Width
    Image1.Height = Screen.Height
    Image1.Stretch = True
    
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub
