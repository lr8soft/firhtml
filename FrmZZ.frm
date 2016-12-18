VERSION 5.00
Begin VB.Form FrmZZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "赞助我们"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3420
   Icon            =   "FrmZZ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3420
   StartUpPosition =   3  '窗口缺省
   Begin 工程1.jcbutton jcbutton1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Caption         =   "安装2345好压支持我们"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin 工程1.jcbutton jcbutton2 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Caption         =   "浏览网页支持我们"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FrmZZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
Shell "cmd /c start http://jifendownload.2345.cn/jifen_2345/2345haozip_kfirsoft.exe"
End Sub

Private Sub jcbutton2_Click()
Shell "cmd /c start http://www.2345.com/?kfirsoft"
End Sub
