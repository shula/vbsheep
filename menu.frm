VERSION 5.00
Begin VB.Form Menu 
   ClientHeight    =   2970
   ClientLeft      =   75
   ClientTop       =   75
   ClientWidth     =   2625
   ControlBox      =   0   'False
   Icon            =   "menu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Kill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get &Real"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sound &Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Sleep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   1920
      Picture         =   "menu.frx":0CFA
      Top             =   2280
      Width           =   630
   End
   Begin VB.Image Awake 
      Height          =   600
      Left            =   1920
      Picture         =   "menu.frx":146C
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Asleep 
      Height          =   600
      Left            =   1920
      Picture         =   "menu.frx":167D
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Mute_Image 
      Height          =   300
      Left            =   2040
      Picture         =   "menu.frx":1BB9
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Loud_Image 
      Height          =   300
      Left            =   2040
      Picture         =   "menu.frx":1F62
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Last_Control As Control
Public Normal_Background As Long
Public Normal_Foreground As Long


Private Sub Command2_Click()
    BG.Show
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLight Command2
    
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLight Command3
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLight Command4
End Sub

Private Sub Form_Load()
    Set Last_Control = Command1
    Normal_Background = Command1.BackColor
    'Normal_Foreground = Command1.ForeColor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub Form_LostFocus()
    Me.Hide
End Sub

Public Sub RePosition(X, Y)
    Dim Margin As Long, W As Long, H As Long
    
    Margin = 250
    W = Me.Width + Margin
    H = Me.Height + Margin * 2
    
    If Me.Left + W > Screen.Width Then Me.Left = X - W Else Me.Left = X
    If Me.Top + H > Screen.Height Then Me.Top = Y - H Else Me.Top = Y
    
End Sub



Sub HighLight(ByRef c As Control)
    Last_Control.BackColor = Normal_Background
    'Last_Control.foreColor = Normal_Foreground
    
    c.BackColor = vbYellow
    'c.ForeColor = vbWhite
    Set Last_Control = c
End Sub




Private Sub Command1_Click()
    If Sleep_Countdown = 0 Then Go_To_Sleep Else Wake_Up
End Sub

Public Sub Go_To_Sleep()
    Sleep_Countdown = 60 * 60  '===>    60 seconds * 60 mintes = 1 hour
    Asleep.Visible = True
    Awake.Visible = False
End Sub

Public Sub Wake_Up()
    Sleep_Countdown = 0
    Direction = ""
    Asleep.Visible = False
    Awake.Visible = True
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLight Command1
End Sub

Private Sub Command3_Click()
    Mute_Mode = Not Mute_Mode
    Mute_Image.Visible = Mute_Mode
    Loud_Image.Visible = Not Mute_Mode
End Sub

Private Sub Command4_Click()
    Unload Me
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Last_Control.BackColor = Normal_Background
    'Last_Control.ForeColor = Normal_Foreground
End Sub
