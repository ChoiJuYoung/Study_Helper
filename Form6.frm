VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   Caption         =   "½ºÅé¿öÄ¡"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   1920
   ScaleWidth      =   7590
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.jcbutton jcbutton2 
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1440
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
      Caption         =   "Stop"
      ForeColor       =   16777215
      ForeColorHover  =   16761087
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   2040
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
      TabIndex        =   4
      Top             =   0
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
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblSec 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "00"
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
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "00"
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
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "00"
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
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sec As Integer, Min As Integer, Hou As Integer
Public Time1 As Long, Time2 As Long, Time3 As Long


Private Sub Form_Load()
Sec = 0
Min = 0
Hou = 0
Time3 = 0
End Sub

Private Sub jcbutton1_Click()
Time1 = GetTickCount + Time3
Timer1.Enabled = True
End Sub

Private Sub jcbutton2_Click()
Time3 = Val(Time1) - Val(Time2)
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Time2 = GetTickCount

Sec = Int((Time2 - Time1) / 1000)
If Val(Sec) >= 60 Then
    Sec = Val(Sec) - 60
    Min = Val(Min) + 1
End If
If Val(Min) >= 60 Then
    Min = Val(Min) - 60
    Hou = Val(Hou) + 1
End If

lblTimer = Hou
lblMin = Min
lblSec = Sec

End Sub
