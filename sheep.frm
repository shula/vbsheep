VERSION 5.00
Begin VB.Form SheepForm 
   BackColor       =   &H00000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   FillColor       =   &H00800000&
   Icon            =   "sheep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Debug"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Walk 
      Interval        =   90
      Left            =   0
      Top             =   1200
   End
   Begin VB.Timer Anim_timer 
      Enabled         =   0   'False
      Left            =   960
      Top             =   1200
   End
   Begin VB.Timer SlowPacer 
      Interval        =   1000
      Left            =   480
      Top             =   1200
   End
   Begin Transparent.ucAniGIF GIF 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   600
      _extentx        =   1058
      _extenty        =   1058
      gif             =   "sheep.frx":0CFA
   End
End
Attribute VB_Name = "SheepForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurX As Double
Private CurY As Double
Private Resource_IDs As Dictionary


Public Sub Change_Animation(What)       ' ------------ GIF resource -------------

    What = LCase(What)
    Select Case What
        Case "meteor"
            Anim_Speed = MEDIUM
            Direction = "bottom left"
            Walk_Steps = 50
            Play_Sound "OHVERYNICE"
        Case "pee"
            Direction = "stop"
            Play_Sound "RUNNINGWATER"
        Case "climb-from-left"
            Direction = "climb"
            Play_Sound "SHEEP3"
        Case "black"
            Direction = "left"
            Walk_Steps = 100
            Play_Sound "SHEEP1"
        Case "z-axis"
            Direction = "left"
            Walk_Steps = 100
            Play_Sound "FALLING"
        Case Else:
            MsgBox "programmer error, asked for gif resource ID=" + What, vbExclamation
            Exit Sub
    End Select
    
    Anim_timer.Enabled = False
    If What <> Current_GIF Then
        Load_resouce_as_GIF Resource_IDs(What)
        Current_GIF = What
    End If
    
End Sub

Public Sub anim(What)               ' ---------- frame by frame -------------
    Dim Anim_Speed As Long
    
    
    Me.SlowPacer.Enabled = False
    Select Case What
        Case "meteor"
            Anim_From = 101
            Anim_To = 116
            Anim_Speed = MEDIUM
            Direction = "bottom left"
            Walk_Steps = 50
            Play_Sound "OHVERYNICE"
        Case "piss"
            Anim_From = 120
            Anim_To = 123
            Anim_Speed = SLOW
            Direction = "stop"
            Play_Sound "RUNNINGWATER"
        Case "climb-from-left"
            Anim_From = 140
            Anim_To = 143
            Anim_Speed = MEDIUM
            Direction = "climb"
            Play_Sound "SHEEP3"
        Case "black"
            Anim_From = 130
            Anim_To = 133
            Anim_Speed = MEDIUM
            Direction = "left"
            Walk_Steps = 100
            Play_Sound "SHEEP1"
        Case Else:
            Me.SlowPacer.Enabled = True
            Exit Sub
    End Select
    
    Anim_timer.Interval = Anim_Speed
    Anim_timer.Enabled = True
    
End Sub

Private Sub Anim_timer_Timer()
'    Me.Picture = LoadResPicture(Anim_From, vbResBitmap)
    Anim_From = Anim_From + 1
    
    If Anim_From > Anim_To Then
        Anim_timer.Enabled = False
        Me.SlowPacer.Enabled = True
    End If
    
End Sub



Private Sub Command2_Click()
            
    'Gif.LoadAnimatedGIF_Array gifs(0)
    Load_resouce_as_GIF 300
    
    'If Not InIDE Then PlaySound "GREATSUCCESS2", App.hInstance, SND_RESOURCE Or SND_ASYNC
    'anim "climb-from-left"
End Sub





Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27:
            Unload Me
            End
        Case Else:
            'MsgBox KeyAscii
    End Select
            
        
End Sub

Private Sub Play_Sound(id As Variant)
    If Not Mute_Mode Then PlaySound id, App.hInstance, SND_RESOURCE Or SND_ASYNC
End Sub

Private Sub Form_Load()
    Randomize Timer
    If Command = "debug" Then Me.Command2.Visible = True
   
    Set Resource_IDs = New Dictionary
    Resource_IDs("black") = 300
    Resource_IDs("pee") = 301
    Resource_IDs("run left") = 302
    Resource_IDs("z-axis") = 303
    Resource_IDs("BLACK") = 300
    Resource_IDs("PEE") = 301
    Resource_IDs("RUN LEFT") = 302
    Resource_IDs("Z-AXIS") = 303
   
    
    Seconds = 0
    Sleep_Countdown = 0
    Mute_Mode = False
    Step_Size = Default_Step_Size
    Screen_Width = Twips_To_Pixels_x(Screen.Width)
    Screen_height = Twips_To_Pixels_y(Screen.Height)
        
    Me.Left = GetSetting(App.EXEName, "position", "left", Screen.Width * 0.8)
    Me.Top = GetSetting(App.EXEName, "position", "top", Screen.Height * 0.8)

    
    ReDim Directions(10) As String
    Directions(0) = "left"
    Directions(1) = "right"
    Directions(2) = "bounce-up"
    Directions(3) = "fall"
    Directions(4) = "climb-from-left"
    Directions(5) = "walkdown"
    Walk_Steps = Int(Rnd * 60 + 10)
   
    
    Dim I As Integer
    'Ex: all transparent at ratio 140/255
    'ActiveTransparency Me, True, False, 140, Me.BackColor
    'Ex: Form transparent, visible component at ratio 140/255
    'ActiveTransparency Me, True, True, 140, Me.BackColor
    'Example display the form transparency degradation
    
    'ActiveTransparency Me, True, False, 0
    'ActiveTransparency Me, True, False, 255
    ActiveTransparency Me, True, False, 255, vbBlue
    SetFormPosition Me.hWnd, True
    
    
End Sub

Sub Katam()
    Dim I As Integer
    For I = 255 To 0 Step -3   'faster
        ActiveTransparency Me, True, False, I
        Me.Refresh
    Next I
    
End Sub



Public Sub Next_Step()
    Dim X As Long
    Dim Y As Long
    
    If Direction = "stop" Then Exit Sub
    
    Walk_Steps = Walk_Steps - 1
    If Walk_Steps <= 0 Then
        'Walk.Enabled = False
        Direction = ""
        Exit Sub
    End If
    
    
    
    X = Twips_To_Pixels_x(Me.Left)
    Y = Twips_To_Pixels_y(Me.Top)
    
    Select Case Direction
        Case "climb"
            Y = Y - Step_Size
        Case "walkdown"
            Y = Y + Step_Size
        Case "bounce-up"
            Step_Size = Round(Step_Size * (1 / Gravity))
            Y = Y - Step_Size
        Case "fall"
            Step_Size = Round(Step_Size * Gravity)
            Y = Y + Step_Size
            Change_Animation "z-axis"
        Case "left"
            X = X - Step_Size
        Case "bottom left"
            Step_Size = Round(Step_Size * Gravity)
            If Y <= Screen_height Then Y = Y + Step_Size
            If X >= 50 Then X = X - Step_Size
        Case "left"
            X = X - Step_Size
        Case "right"
            X = X + Step_Size
        Case Else
            Exit Sub
    End Select
    
    Dim Dont_Move As Boolean
    Dont_Move = False
    Margin = 5   'pixels
    
    If Y + Sheep_Size > Screen_height Then
        Direction = "bounce-up"
        Y = Screen_height - Sheep_Size - Margin
        Play_Sound "bounce"
        Stop_GIF_animation
        Dont_Move = True
    End If
    If X < 2 Then
        Direction = "right"
        X = 1
        Dont_Move = True
    End If
    If Y < 2 Then
        Direction = "fall"
        Y = 1
        Dont_Move = True
        Change_Animation "z-axis"
        Play_Sound "falling"
    End If
    If X > Screen_Width Then
        Direction = "climb"
        X = Screen_Width - Sheep_Size
        anim "climb-from-left"
        Dont_Move = True
    End If
    'If Dont_Move Then Exit Sub
    
    'Me.Move Me.Left - Pixels_to_Twips_x(x), Me.Top - Pixels_to_Twips_y(y)
    Me.Left = Pixels_to_Twips_x(X)
    Me.Top = Pixels_to_Twips_y(Y)
    Save
    
End Sub

Public Sub Save()
    SaveSetting App.EXEName, "position", "left", Me.Left
    SaveSetting App.EXEName, "position", "top", Me.Top
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Save
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' --- grade
    'Katam
End Sub



Private Sub Gif_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'right click
    If Button = 2 Then
        Menu.RePosition Me.Left + X, Me.Top + Y
        Menu.Show
    End If
    
    If Button = 1 Then Save

End Sub

Private Sub SlowPacer_Timer()
    Dim Ret As Long
    Seconds = Seconds + 1
    If Seconds > 2147483645 Then Seconds = 0
    If Sleep_Countdown > 0 Then Sleep_Countdown = Sleep_Countdown - 1
     
    If Int(Rnd * 10) = 1 Then anim "black"
    If Int(Rnd * 10) = 2 Then anim "piss"
    If Int(Rnd * 10) = 3 Then anim "meteor"
    
End Sub

Private Sub ucAniGIF1_Click()

End Sub

Private Sub Walk_Timer()
    
    If Sleep_Countdown > 0 Then Exit Sub
    If Direction = "stop" Then Exit Sub
    
    If IsEmpty(Direction) Then Direction = ""
    If Direction = "" Then
        ' -------- init -------
        z = Int(Rnd * 10)
        Step_Size = Default_Step_Size
        If z >= 4 Then Exit Sub
        Direction = Directions(z)
        If Direction = "fall" Then Load_resouce_as_GIF Int(Resource_IDs("z-axis"))
        'Walk.Enabled = True
        Walk_Steps = Int(Rnd * 70) + 10
    Else
        Next_Step
    End If
        
End Sub




Private Sub Stop_GIF_animation()
    GIF.Action = gfaPause
End Sub

Private Sub Load_resouce_as_GIF(Resource_ID As Integer)
    Dim Data() As Byte
    Data = LoadResData(Resource_ID, "CUSTOM")
    YesNo = GIF.LoadAnimatedGIF_Array(Data)
    GIF.Action = gfaPlay
    GIF.Stretch = gfsActualSize
    
    'Static TempFile As String
    'If TempFile = "" Then TempFile = CreateTempFilename_With_Extension("", "sheep-", ".gif")
    'LoadDataIntoFile Resource_ID, TempFile
    'Gif.LoadAnimatedGIF_File TempFile
End Sub

Public Sub LoadDataIntoFile(DataName As Integer, FileName As String)
    Dim myArray() As Byte
    Dim myFile As Long
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
        Put #myFile, , myArray
        Close #myFile
    End If
End Sub









' ---------------------------- move, drag ---------------------------------
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurX = X
    CurY = Y
End Sub
Private Sub Gif_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurX = X
    CurY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'And Shift = 1
    If Button = 1 Then 'use mouse_left_button
        Direction = "stop"
        Step_Size = Default_Step_Size
        Me.Move Me.Left + (X - CurX), Me.Top + (Y - CurY)
    End If
End Sub
Private Sub Gif_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'And Shift = 1
    If Button = 1 Then 'use mouse_left_button
       'Me.Move Me.Left + (X - CurX), Me.Top + (Y - CurY)
        Step_Size = Default_Step_Size
        Direction = "stop"
        Me.Move Me.Left + GIF.Left + (X - CurX), Me.Top + GIF.Top + (Y - CurY)
    End If
End Sub

