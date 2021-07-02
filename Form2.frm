VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Timer"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2175
   ScaleWidth      =   7680
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   2160
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Start"
      ForeColor       =   16777215
      ForeColorHover  =   16761087
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "59"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblSec 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "59"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Time1 As Long, Time2 As Long, Time3 As Long
Public Sec As Long, Min As Long, Hour As Long

Private Sub jcbutton1_Click()
Time1 = InputBox("½Ã°£À» ÀÔ·ÂÇØ ÁÖ½Ê½Ã¿À.(ÃÊ´ÜÀ§)")
Time2 = GetTickCount + Time1 * 1000
Time3 = GetTickCount
Sec = Int((Time2 - Time3) / 1000)

Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Time3 = GetTickCount
Min = 0
Hour = 0
Sec = Int((Time2 - Time3) / 1000)

Do Until Sec < 60
    If Sec >= 60 Then
        Sec = Sec - 60
        Min = Min + 1
    ElseIf Sec <= -1 Then
        Sec = 59
        Min = Min - 1
    End If
Loop

Do Until Min < 60
    If Min >= 60 Then
        Min = Min - 60
        Hour = Hour + 1
    ElseIf Min <= -1 Then
        Min = 59
        Hour = Hour - 1
    End If
Loop

If Hour <= -1 Then
    Sec = 0
    Min = 0
    Hour = 0
    MsgBox "Alarm."
    Timer1.Enabled = False
End If

lblSec = Sec
lblMin = Min
lblTimer = Hour
End Sub

