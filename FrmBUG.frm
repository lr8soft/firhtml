VERSION 5.00
Begin VB.Form FrmBUG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提交BUG"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5475
   Icon            =   "FrmBUG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5475
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "183813847@qq.com"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "发送邮件到(点击复制)："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2550
   End
End
Attribute VB_Name = "FrmBUG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
Clipboard.Clear
Clipboard.SetText Label2.Caption
MsgBox "文字已复制到剪贴板！", vbInformation
End Sub
