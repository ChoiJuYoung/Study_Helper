VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "WordPad"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9240
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   6975
   ScaleWidth      =   9240
   StartUpPosition =   2  '화면 가운데
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   9600
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.jcbutton jcbutton2 
      Height          =   495
      Left            =   12360
      TabIndex        =   2
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "SavE"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   495
      Left            =   12360
      TabIndex        =   1
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Open"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.TextBox Text1 
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
      Height          =   6975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Menu 메뉴 
      Caption         =   "메뉴"
      Begin VB.Menu MenuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuSave 
         Caption         =   "SavE"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
Dialog1.Filter = "Text Files (*.txt)"
On Error GoTo errorS1
Dim FF As Integer, Temp As String, 경로 As String
Dialog1.ShowOpen
경로 = Dialog1.FileName
FF = FreeFile()
Open 경로 For Input As #FF
Do Until EOF(FF)
Line Input #FF, Temp
Text1 = Text1 & Temp & vbCrLf
Loop
Close #FF
Exit Sub

errorS1:
Exit Sub
End Sub
Private Sub jcbutton2_Click()
Dialog1.Filter = "Text Files (*.txt)"
On Error GoTo errorS2
Dim 경로 As String
Dialog1.ShowSave
경로 = Dialog1.FileName
Open 경로 & ".txt" For Output As #1
Print #1, Text1
Close #1
Exit Sub

errorS2:
Exit Sub
End Sub

Private Sub MenuOpen_Click()
jcbutton1_Click
End Sub

Private Sub MenuSave_Click()
jcbutton2_Click
End Sub
