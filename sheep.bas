Attribute VB_Name = "SheepMod"


Public Const DEFAULT_TIME_TO_CLOSE = 3
Public Const MAXINT = 65535
Public Const SLOW = 350
Public Const MEDIUM = 180
Public Const FAST = 90
Public Const Gravity = 1.18
Public Const Default_Step_Size = 3
Public Const Sheep_Size = 41
Public Anim_From As Long, Anim_To As Long
Public Step_Size As Integer
Public Screen_height As Integer
Public Screen_Width As Integer
Public Walk_Steps As Integer
Public Direction As String
Public Directions As Variant
Public Seconds As Long
Public Sleep_Countdown As Long
Public Mute_Mode As Boolean

Public Current_GIF As String
