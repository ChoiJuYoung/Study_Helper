VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "CalCulator"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3015
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2400
   ScaleWidth      =   3015
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.jcbutton CmdFac 
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "Fac"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdTan 
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "Tan"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdCos 
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "Cos"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdSin 
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "Sin"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdEqual 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "="
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdDivi 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "/"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdMulti 
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "X"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdMinus 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "-"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdPlus 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "+"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdSqr 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "Sqr"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdC 
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "C"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton CmdBack 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "¡ç"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd00 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "0"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd03 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "3"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd02 
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "2"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd01 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "1"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd06 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "6"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd05 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "5"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd04 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "4"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd09 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "9"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd08 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "8"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton Cmd07 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   3
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
      Caption         =   "7"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   8421504
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FirstValue As String
Public Mode As String
Public Rad As String
Private Const pi As Double = 3.14159265358979
Private Const j As Byte = 14

Private Function Fac(num As Long) As Variant
Dim i As Long
Fac = 1
    For i = 1 To num
        Fac = i * Fac
    Next
End Function
Private Function xsin(X As Double, ny As Byte) As Double
Dim i As Long, Y As Variant, t As Boolean, q As Double
    Y = IIf(ny Mod 2 = 0, ny - 1, ny)
    For i = 3 To Y Step 2
        If t = False Then
        If i = 3 Then q = X
            q = q - (X ^ i) / Fac(i)
        ElseIf t = True Then
            q = q + (X ^ i) / Fac(i)
        End If
'-----------------------------------------------------
'-----------------------------------------------------
        If t = False Then
            t = True
        ElseIf t = True Then
            t = False
        End If
    Next i
xsin = q
End Function

Private Function xcos(X As Double, ny As Byte) As Double
Dim i As Long, Y As Variant, t As Boolean, q As Double
    Y = IIf(ny Mod 2 = 0, ny, ny - 1)
    For i = 2 To Y Step 2
        If t = False Then
        If i = 2 Then q = 1
            q = q - (X ^ i) / Fac(i)
        ElseIf t = True Then
            q = q + (X ^ i) / Fac(i)
        End If
'-----------------------------------------------------
'-----------------------------------------------------
        If t = False Then
            t = True
        ElseIf t = True Then
            t = False
        End If
    Next i
xcos = q
End Function

Private Sub Cmd00_Click()
TxtValue = Val(TxtValue) * 10
End Sub

Private Sub Cmd01_Click()
TxtValue = Val(TxtValue) * 10 + 1
End Sub

Private Sub Cmd02_Click()
TxtValue = Val(TxtValue) * 10 + 2
End Sub

Private Sub Cmd03_Click()
TxtValue = Val(TxtValue) * 10 + 3
End Sub

Private Sub Cmd04_Click()
TxtValue = Val(TxtValue) * 10 + 4
End Sub

Private Sub Cmd05_Click()
TxtValue = Val(TxtValue) * 10 + 5
End Sub

Private Sub Cmd06_Click()
TxtValue = Val(TxtValue) * 10 + 6
End Sub

Private Sub Cmd07_Click()
TxtValue = Val(TxtValue) * 10 + 7
End Sub

Private Sub Cmd08_Click()
TxtValue = Val(TxtValue) * 10 + 8
End Sub

Private Sub Cmd09_Click()
TxtValue = Val(TxtValue) * 10 + 9
End Sub

Private Sub CmdBack_Click()
TxtValue = Int(Val(TxtValue) / 10)
End Sub

Private Sub CmdC_Click()
TxtValue = 0
End Sub

Private Sub CmdDivi_Click()
Mode = "Divide"
FirstValue = Val(TxtValue)
TxtValue = ""
End Sub

Private Sub CmdEqual_Click()
If Mode = "Plus" Then
    TxtValue = Val(TxtValue) + Val(FirstValue)
ElseIf Mode = "Minus" Then
    TxtValue = Val(FirstValue) - Val(TxtValue)
ElseIf Mode = "Multiple" Then
    TxtValue = Val(FirstValue) * Val(TxtValue)
ElseIf Mode = "Divide" Then
    TxtValue = Val(FirstValue) / Val(TxtValue)
End If
Mode = ""
FirstValue = 0

End Sub

Private Sub CmdFac_Click()
TxtValue = Fac(TxtValue)
End Sub

Private Sub CmdMinus_Click()
Mode = "Minus"
FirstValue = Val(TxtValue)
TxtValue = ""
End Sub

Private Sub CmdMulti_Click()
Mode = "Multiple"
FirstValue = Val(TxtValue)
TxtValue = ""
End Sub

Private Sub CmdPlus_Click()
Mode = "Plus"
FirstValue = Val(TxtValue)
TxtValue = ""
End Sub

Private Sub CmdSqr_Click()
TxtValue = Sqr(TxtValue)
End Sub

Private Sub Form_Load()
Mode = ""
FirstValue = 0
End Sub

Private Sub cmdCos_Click()
Dim X As Double, C As Double
    X = CDbl(Rad)
    C = xcos(X, 170)
    C = Round(C, j)
    TxtValue = C
End Sub

Private Sub CmdSin_Click()
Dim X As Double, s As Double
    X = CDbl(Rad)
    s = xsin(X, 170)
    s = Round(s, j)
TxtValue = s
End Sub

Private Sub cmdTan_Click()
Dim X As Double, s As Double, C As Double, t As Double, ta As String
    X = CDbl(Rad)
    s = xsin(X, 170)
    C = xcos(X, 170)
    s = Round(s, j)
    C = Round(C, j)
'-----------------------------------------------------
'-----------------------------------------------------
    If C > 0 Then
        t = s / C
    End If
    ta = IIf(C = 0, "¡Ä", t)
    TxtValue = ta
End Sub

Private Sub TxtValue_Change()
Rad = CDbl(CDbl(Val(TxtValue))) * (pi / 180)
End Sub
