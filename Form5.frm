VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   Caption         =   "StudyHelper"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   1560
      Picture         =   "Form5.frx":08CA
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Press Any Key"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Study Helper. Made By. "
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   1080
      Y1              =   1560
      Y2              =   2160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   720
      X2              =   1080
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   720
      X2              =   360
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1320
      X2              =   360
      Y1              =   1560
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   1320
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ABC As New WinHttpRequest
Public StarDire As Integer, StarVal As Integer, StarColor As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If StarDire = 6 Then
    Form1.Show
    Unload Me
End If

If Len(Dir("C:\Windows\system32\COMDLG32.ocx")) = 0 Then
    FileCopy App.Path & "\COMDLG32.ocx", "C:\Windows\system32\COMDLG32.ocx"
End If
End Sub

Private Sub Form_Load()
StarDire = 1
StarVal = 1
StarColor = 1
Line3.Visible = False
Line2.Visible = False
Line4.Visible = False
Line5.Visible = False
Line1.X2 = 120
Line2.X2 = 1320
Line2.Y2 = 1560
Line3.X1 = 360
Line3.Y1 = 2160
Line4.X2 = 720
Line4.Y2 = 1200
Line5.X1 = 1080
Line5.Y1 = 2160
Label1.Top = -510
Label2.Top = -255

On Error GoTo ErOr:
Dim Name As String
ABC.Open "GET", "http://pos.woobi.co.kr/xe/Board"
ABC.Send
Name = Split(Split(ABC.ResponseText, "<td class=" & """" & "author" & """" & "><div class=" & """" & "member_224" & """" & ">")(1), "</div>")(0)

Label1 = Label1 & Name

Exit Sub
ErOr:
Label1 = Label1 & "최주영"

Exit Sub
End Sub

Private Sub Timer1_Timer()
If StarDire = 1 Then
    If Val(StarVal) <= 99 Then
        Line1.X2 = 120 + Val(1200) * Val(StarVal) / 100
        StarVal = Val(StarVal) + 1
    Else
        StarVal = 0
        StarDire = 2
        Line2.Visible = True
    End If
ElseIf StarDire = 2 Then
    If Val(StarVal) <= 99 Then
        Line2.X2 = 1320 - Val(960) * Val(StarVal) / 100
        Line2.Y2 = 1560 + Val(600) * Val(StarVal) / 100
        StarVal = Val(StarVal) + 1
    Else
        StarVal = 0
        StarDire = 3
        Line3.Visible = True
    End If
ElseIf StarDire = 3 Then
    If Val(StarVal) <= 99 Then
        Line3.X1 = 360 + Val(360) * Val(StarVal) / 100
        Line3.Y1 = 2160 - Val(960) * Val(StarVal) / 100
        StarVal = Val(StarVal) + 1
    Else
        StarVal = 0
        StarDire = 4
        Line4.Visible = True
    End If
ElseIf StarDire = 4 Then
    If Val(StarVal) <= 99 Then
        Line4.X2 = 720 + Val(360) * Val(StarVal) / 100
        Line4.Y2 = 1200 + Val(960) * Val(StarVal) / 100
        StarVal = Val(StarVal) + 1
    Else
        StarVal = 0
        StarDire = 5
        Line5.Visible = True
    End If
ElseIf StarDire = 5 Then
    If Val(StarVal) <= 99 Then
        Line5.X1 = 1080 - Val(960) * Val(StarVal) / 100
        Line5.Y1 = 2160 - Val(600) * Val(StarVal) / 100
        StarVal = Val(StarVal) + 1
    Else
        StarVal = 0
        StarDire = 6
    End If
Else
    If Val(StarVal) <= 99 Then
        Label1.Top = -510 + Val(1935) * Val(StarVal) / 100
        Label2.Top = -255 + Val(1935) * Val(StarVal) / 100
        StarVal = Val(StarVal) + 1
    Else
        Timer1.Interval = 500
        If StarColor = 1 Then
            Line1.BorderColor = RGB(255, 255, 255)
            Line2.BorderColor = RGB(255, 255, 255)
            Line3.BorderColor = RGB(255, 255, 255)
            Line4.BorderColor = RGB(255, 255, 255)
            Line5.BorderColor = RGB(255, 255, 255)
            StarColor = 2
        ElseIf StarColor = 2 Then
            Line1.BorderColor = RGB(255, 255, 0)
            Line2.BorderColor = RGB(255, 255, 0)
            Line3.BorderColor = RGB(255, 255, 0)
            Line4.BorderColor = RGB(255, 255, 0)
            Line5.BorderColor = RGB(255, 255, 0)
            StarColor = 3
        ElseIf StarColor = 3 Then
            Line1.BorderColor = RGB(255, 0, 255)
            Line2.BorderColor = RGB(255, 0, 255)
            Line3.BorderColor = RGB(255, 0, 255)
            Line4.BorderColor = RGB(255, 0, 255)
            Line5.BorderColor = RGB(255, 0, 255)
            StarColor = 4
        ElseIf StarColor = 4 Then
            Line1.BorderColor = RGB(0, 255, 255)
            Line2.BorderColor = RGB(0, 255, 255)
            Line3.BorderColor = RGB(0, 255, 255)
            Line4.BorderColor = RGB(0, 255, 255)
            Line5.BorderColor = RGB(0, 255, 255)
            StarColor = 5
        ElseIf StarColor = 5 Then
            Line1.BorderColor = RGB(255, 0, 0)
            Line2.BorderColor = RGB(255, 0, 0)
            Line3.BorderColor = RGB(255, 0, 0)
            Line4.BorderColor = RGB(255, 0, 0)
            Line5.BorderColor = RGB(255, 0, 0)
            StarColor = 6
        ElseIf StarColor = 6 Then
            Line1.BorderColor = RGB(0, 255, 0)
            Line2.BorderColor = RGB(0, 255, 0)
            Line3.BorderColor = RGB(0, 255, 0)
            Line4.BorderColor = RGB(0, 255, 0)
            Line5.BorderColor = RGB(0, 255, 0)
            StarColor = 7
        ElseIf StarColor = 7 Then
            Line1.BorderColor = RGB(0, 0, 255)
            Line2.BorderColor = RGB(0, 0, 255)
            Line3.BorderColor = RGB(0, 0, 255)
            Line4.BorderColor = RGB(0, 0, 255)
            Line5.BorderColor = RGB(0, 0, 255)
            StarColor = 1
        End If
    End If
End If
End Sub
