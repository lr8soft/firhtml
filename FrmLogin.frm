VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firsoft用户登录"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3960
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin 工程1.jcbutton jcbutton1 
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "登录"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin 工程1.jcbutton jcbutton2 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "注册"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "警告:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密 码 :"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginErrorCount As Long '登录错误次数
Private Sub jcbutton1_Click()
    Dim uandpass$, rtn$, codeseed
If Text1.Text = "" Or Text2.Text = "" Then
   Label3.Visible = True
   Label3.Caption = "警告:" & "用户名或密码不得为空！"
   Timer1.Enabled = True
Else
    codeseed = IIf(getHtmlCode(strServerName & "CheckConnect.html") = "网络已连接上", _
               getHtmlCode(strServerName & "Show_UsageCode.asp"), _
               "Error")
  ' MsgBox codeseed
    If codeseed = "Error" Or codeseed = "OutTime" Then
        MsgBox "登录超时，请检查网络设置！", vbExclamation, "无法连接到服务器"
        Exit Sub
    End If
   ' MsgBox codeseed
   ' MsgBox PassCodeC(codeseed)
   
   
    uandpass = strServerName & "login.asp?username=" & Text1.Text & "&password=" & Text2.Text
    rtn = getHtmlCode(uandpass)
    'MsgBox rtn
    Label3.Visible = True
    Label3.Caption = "登录中..."
    Select Case LCase$(rtn)
        Case "okay"
        MsgBox "登录成功！", vbInformation
        Label3.Visible = False
        username = Text1.Text
        password = Text2.Text
        areyoulogin = True
        FrmMain.wel_exit.Visible = True
        FrmMain.wel_exit.Caption = "欢迎！" & username & " 点击退出"
        FrmMain.wel_exit2.Visible = True
        FrmMain.wel_exit2.Caption = "欢迎！" & username & " 点击退出"
       ' MsgBox gamefrm.passwd & Chr(13) & gamefrm.strServerName
        Unload Me
        Case "no"
        MsgBox "用户名不存在或密码错误！", vbExclamation, "用户名或密码错误"
        Label3.Visible = False
        Case "outtime"
        MsgBox "登录超时，请检查网络设置或重试！", vbExclamation, "无法连接到服务器"
        Label3.Visible = False
        Case Else
        MsgBox "未知的登录错误...", vbCritical, "错误"
        Label3.Visible = False
    End Select


End If
End Sub

Private Sub jcbutton2_Click()
    Dim uandpass$, rtn$, codeseed
If Text1.Text = "" Or Text2.Text = "" Then
   Label3.Visible = True
   Label3.Caption = "警告:" & "用户名或密码不得为空！"
   Timer1.Enabled = True
Else
   If Len(Text2.Text) > 20 Or Len(Text2.Text) < 6 Then
        MsgBox "密码长度必须大于6位，小于20位！", vbExclamation
        Exit Sub
    End If

    codeseed = IIf(getHtmlCode(strServerName & "CheckConnect.html") = "网络已连接上", _
               getHtmlCode(strServerName & "Show_UsageCode.asp"), _
               "Error")
    'MsgBox codeseed
    If codeseed = "Error" Or codeseed = "OutTime" Then
        MsgBox "注册超时，请检查网络设置！", vbExclamation, "无法连接到服务器"

        Exit Sub
    End If
    'MsgBox codeseed
    'MsgBox PassCodeC(codeseed)
    uandpass = strServerName & "reg.asp?username=" & Text1.Text & "&password=" & Text2.Text & "&qq=" & "10086"
    rtn = getHtmlCode(uandpass)
    'MsgBox rtn
    Select Case LCase$(rtn)
        Case "added"
        MsgBox "注册成功！", vbInformation, "OK"
        Unload Me
        Case "username existed"
        MsgBox "用户名重复！", vbExclamation, "用户名重复"
        Case "null password length"
        MsgBox "密码长度必须不少于6位！", vbExclamation, "密码长度"
        Case "outtime"
        MsgBox "注册超时，请检查网络设置！", vbExclamation, "无法连接到服务器"
        Case Else
        MsgBox "未知的错误...", vbCritical, "错误"
    End Select


End If
End Sub

Private Sub Timer1_Timer()
Label3.Visible = False
Timer1.Enabled = False
End Sub
