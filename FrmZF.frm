VERSION 5.00
Begin VB.Form FrmZF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "工程字符编码"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   Icon            =   "FrmZF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5385
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option2 
      Caption         =   "GB2312"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "UTF-8"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin 工程1.jcbutton jcbutton1 
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "确定"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin 工程1.jcbutton jcbutton2 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "取消"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FrmZF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Private Sub Label2_Click()
End Sub

Private Sub jcbutton1_Click()
If a <> "" Then
ZF = a
MsgBox "成功将字符编码调整为" & a, vbInformation
ZFGS = "<head>" & Chr(13) & "<meta http-equiv=" & """" & "Content-Type" & """" & " content=" & """" & "text/html; charset=" & ZF & """" & "/>" & Chr(13) & "</head>"
Unload Me
End If
End Sub

Private Sub jcbutton2_Click()
Unload Me
End Sub

Private Sub Option1_Click()
a = "UTF-8"
End Sub

Private Sub Option2_Click()
a = "GB2312"
End Sub
