VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Main Menu"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   9300
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.jcbutton jcbutton6 
      Height          =   615
      Left            =   6840
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   10
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
      Caption         =   "Åë¿ª±â"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton5 
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   10
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
      Caption         =   "¼ö¿­ÀÇ ÇÕ"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton4 
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   10
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
      Caption         =   "Stop Watch"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton3 
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   10
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
      Caption         =   "WordPad"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton2 
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   10
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
      Caption         =   "Calculator"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   10
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
      Caption         =   "Timer"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":08CA
      Top             =   2520
      Width           =   6720
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   840
      Picture         =   "Form1.frx":28AF
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
Form2.Show
End Sub

Private Sub jcbutton2_Click()
Form3.Show
End Sub

Private Sub jcbutton3_Click()
Form4.Show
End Sub

Private Sub jcbutton4_Click()
Form6.Show
End Sub

Private Sub jcbutton5_Click()
Form7.Show
End Sub

Private Sub jcbutton6_Click()
Form8.Show
End Sub
