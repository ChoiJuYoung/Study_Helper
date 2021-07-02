VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "통역기"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   3810
   ScaleWidth      =   5160
   StartUpPosition =   2  '화면 가운데
   Begin Project1.jcbutton jcbutton1 
      Height          =   225
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   397
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "jcbutton"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "번역하기"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "번역하기"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox TxtEng 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Text            =   "Translater"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox TxtKor 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Text            =   "통역기"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblKor 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "eNglisH tO kOreaN"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label lblEng 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "kOreaN tO eNglisH"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ABC As New WinHttpRequest
Private Sub Command1_Click()
ABC.Open "POST", "http://kr.babelfish.yahoo.com/translate_txt"
ABC.SetRequestHeader "Referer", "http://kr.babelfish.yahoo.com/translate_txt"
ABC.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
ABC.Send "fr=bf-res&trtext=" & UTF16(TxtKor) & "&lp=ko_en"
lblEng = Split(Split(ABC.ResponseText, "<div style=""padding:0.6em;"">")(1), "</div>")(0)
End Sub

Private Sub Command2_Click()
ABC.Open "POST", "http://kr.babelfish.yahoo.com/translate_txt"
ABC.SetRequestHeader "Referer", "http://kr.babelfish.yahoo.com/translate_txt"
ABC.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
ABC.Send "fr=bf-res&trtext=" & TxtEng & "&lp=en_ko"
lblKor = Split(Split(ABC.ResponseText, "<div style=""padding:0.6em;"">")(1), "</div>")(0)
End Sub
