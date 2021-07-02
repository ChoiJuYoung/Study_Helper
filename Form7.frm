VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "수열의 합"
   ClientHeight    =   2520
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   3870
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   2520
   ScaleWidth      =   3870
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame3 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox TxtRA 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   720
         TabIndex        =   23
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxtRR 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   720
         TabIndex        =   22
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtRN 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   720
         TabIndex        =   21
         Text            =   "1"
         Top             =   1080
         Width           =   615
      End
      Begin Project1.jcbutton jcbutton3 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   65535
         Caption         =   "계산하기"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "a ="
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
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "r = "
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
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "n = "
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
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "an ="
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
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblrV 
         BackStyle       =   0  '투명
         Caption         =   "Sum : "
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
         Left            =   1560
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblRAN 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "1"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   1440
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin Project1.jcbutton jcbutton2 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   255
         Caption         =   "계산하기"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.TextBox TxtDN 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   720
         TabIndex        =   16
         Text            =   "1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TxtDD 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   720
         TabIndex        =   15
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtDA 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   720
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblDAN 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "1"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbldV 
         BackStyle       =   0  '투명
         Caption         =   "Sum : "
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
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "an ="
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
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "n = "
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
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "d = "
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
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "a ="
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
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin Project1.jcbutton jcbutton1 
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   33023
         Caption         =   "계산하기"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.TextBox TxtA 
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
         Left            =   1680
         TabIndex        =   5
         Text            =   "1"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TxtN 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   840
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtK 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   1200
         TabIndex        =   3
         Text            =   "1"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblV 
         BackStyle       =   0  '투명
         Caption         =   "Sum : 1"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "K ="
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
         Left            =   840
         TabIndex        =   2
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Σ"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   72
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Menu 메뉴 
      Caption         =   "메뉴"
      Begin VB.Menu 등비 
         Caption         =   "등비수열"
      End
      Begin VB.Menu 등차 
         Caption         =   "등차수열"
      End
      Begin VB.Menu 시그마 
         Caption         =   "Σ"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 등비_Click()
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub 등차_Click()
Frame3.Visible = False
Frame1.Visible = False
Frame2.Visible = True
End Sub

Private Sub 시그마_Click()
Frame2.Visible = False
Frame1.Visible = True
Frame3.Visible = False
End Sub

Private Sub jcbutton1_Click()
Dim B(0 To 100) As String
Dim a As String
Dim Y As Long
Dim X As Long
Dim i As Long
Dim Sum As Long
Dim K As Long
Dim C As Long
Dim M As Long
For i = 0 To 100
    B(i) = ""
Next

a = TxtA
Y = TxtN
X = TxtK
i = 0

Do Until InStr(a, ",") = 0
    C = InStr(a, ",")
    B(i) = Left(a, C - 1)
    a = Mid(a, C + 1)
    i = i + 1
Loop

B(i) = a
Do Until B(i) = ""
    i = i + 1
Loop

i = i - 1

For K = X To Y
    For M = 0 To i
        Sum = Sum + B(M) * K ^ (i - M)
    Next
Next

lblV = "Sum : " & Sum

End Sub

Private Sub jcbutton2_Click()
lbldV = "Sum : " & TxtDN * (2 * TxtDA + (TxtDN - 1) * TxtDD) / 2
End Sub

Private Sub jcbutton3_Click()
lblrV = "Sum : " & TxtRA * (TxtRR ^ TxtRN - 1) / (TxtRR - 1)
End Sub

Private Sub TxtdA_Change()
If TxtDA - TxtDD <= 0 Then
    lblDAN = TxtDD & "n " & TxtDA - TxtDD
Else
    lblDAN = TxtDD & "n +" & TxtDA - TxtDD
End If
End Sub

Private Sub TxtdD_Change()
If TxtDA - TxtDD <= 0 Then
    lblDAN = TxtDD & "n " & TxtDA - TxtDD
Else
    lblDAN = TxtDD & "n +" & TxtDA - TxtDD
End If
End Sub

Private Sub TxtRA_Change()
lblRAN = TxtRA & " * " & TxtRR & " ^(n-1)"
End Sub

Private Sub TxtRR_Change()
lblRAN = TxtRA & " * " & TxtRR & " ^(n-1)"
End Sub
