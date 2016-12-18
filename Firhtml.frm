VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EA8DFBAC-9A7C-471B-B6FD-781EA199403D}#1.0#0"; "prjXTab.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Firhtml2 "
   ClientHeight    =   7905
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13305
   Icon            =   "Firhtml.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   13305
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog savefhm 
      Left            =   1080
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Firhtml页面文件(*.fhm)|*.fhm|"
   End
   Begin prjXTab.XTab XTab1 
      Height          =   7575
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13361
      TabCount        =   2
      TabCaption(0)   =   "代码"
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "codebox"
      Tab(0)ContCtrlCap(2)=   "wel_exit"
      TabCaption(1)   =   "预览"
      TabContCtrlCnt(1)=   2
      Tab(1)ContCtrlCap(1)=   "codepreview"
      Tab(1)ContCtrlCap(2)=   "wel_exit2"
      TabTheme        =   2
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      InActiveTabForeColor=   -2147483631
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   -2147483628
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin SHDocVwCtl.WebBrowser codepreview 
         Height          =   6975
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   10695
         ExtentX         =   18865
         ExtentY         =   12303
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin RichTextLib.RichTextBox codebox 
         Height          =   6975
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   12303
         _Version        =   393217
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Firhtml.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label wel_exit2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "欢迎！%username%"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -65865
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label wel_exit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "欢迎！%username%"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   9135
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
   End
   Begin MSComDlg.CommonDialog opensave 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "网页文件(*.html)|*.html|"
   End
   Begin VB.ListBox projectinfo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      ItemData        =   "Firhtml.frx":03B7
      Left            =   120
      List            =   "Firhtml.frx":03C1
      TabIndex        =   2
      Top             =   4200
      Width           =   1875
   End
   Begin VB.ListBox userkj 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      IntegralHeight  =   0   'False
      ItemData        =   "Firhtml.frx":03D5
      Left            =   120
      List            =   "Firhtml.frx":03F7
      TabIndex        =   1
      Top             =   480
      Width           =   1875
   End
   Begin MSComDlg.CommonDialog saveproject 
      Left            =   600
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Firhtml工程文件(*.fhp)|*.fhp|"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "工程属性"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "可用控件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
   Begin VB.Menu mfile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mNewProject 
         Caption         =   "新建工程(&N)..."
      End
      Begin VB.Menu mOpenProject 
         Caption         =   "打开工程(&O)..."
      End
      Begin VB.Menu mg1 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveProject1 
         Caption         =   "保存工程(&V)..."
      End
      Begin VB.Menu mSaveOther 
         Caption         =   "工程另存为(&E)..."
      End
      Begin VB.Menu mMakeHtml 
         Caption         =   "生成网页文件"
      End
      Begin VB.Menu mg2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "退出(&X)..."
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mCut 
         Caption         =   "剪切(&X)"
      End
      Begin VB.Menu mCopy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu mZhantie 
         Caption         =   "粘贴(&V)"
      End
      Begin VB.Menu mG3 
         Caption         =   "-"
      End
      Begin VB.Menu mYJ 
         Caption         =   "我的语句"
      End
   End
   Begin VB.Menu mSet 
      Caption         =   "设置(&S)"
      Begin VB.Menu mMyYUJU 
         Caption         =   "自定义语句"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mTS 
      Caption         =   "调试(&D)"
      Begin VB.Menu mRunInIE 
         Caption         =   "在浏览器中调试(&R)"
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mlogin 
         Caption         =   "用户管理"
      End
      Begin VB.Menu mg4 
         Caption         =   "-"
      End
      Begin VB.Menu mZZ 
         Caption         =   "赞助我们"
      End
      Begin VB.Menu mBUG 
         Caption         =   "提交BUG"
      End
      Begin VB.Menu mAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pjtitle
Public ars As Boolean
Public gg_path
Public gg_name
Dim X1, X2, X3



Private Sub codebox_KeyUp(KeyCode As Integer, Shift As Integer)
S = True

End Sub

Private Sub codebox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mEdit
    End If
End Sub

Private Sub Form_Load()
areyoulogin = False
strServerName = "http://firhtml.3vzhuji.com/"
firhtmlversion = "Build 2600"
notfirhtmlpj = False
pjtitle = "新工程"
FrmMain.Caption = pjtitle & " - Firhtml2 " & firhtmlversion
ZF = "GB2312"
ars = False
codebox.Text = "<!--以下输入" & """" & "<html>" & """" & "后的代码！-->" & Chr(13)
gg_path = "http://jifendownload.2345.cn/jifen_2345/2345haozip_kfirsoft.exe"
ZFGS = "<head>" & Chr(13) & "<meta http-equiv=" & """" & "Content-Type" & """" & " content=" & """" & "text/html; charset=" & ZF & """" & "/>" & Chr(13) & "</head>"
Exit Sub

  
End Sub

Private Sub Form_Resize()
On Error Resume Next
XTab1.Width = Me.Width - 2610
XTab1.Height = Me.Height - 1215
codebox.Height = XTab1.Height - 600
codebox.Width = XTab1.Width - 240
codepreview.Height = XTab1.Height - 600
codepreview.Width = XTab1.Width - 240

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\bin.html"
Unload FrmMakeMoney
End
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub guanggao_Click()

End Sub

Private Sub mAbout_Click()
frmAbout.Show
End Sub

Private Sub Picture1_DblClick()
codebox.Visible = True
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub mBUG_Click()
FrmBUG.Show
End Sub

Private Sub mCopy_Click()
Clipboard.SetText codebox.SelText
End Sub

Private Sub mCut_Click()
On Error Resume Next
Clipboard.SetText codebox.SelText
    codebox.SelText = ""
End Sub

Private Sub mExit_Click()
End
End Sub

Private Sub mMakeHtml_Click()
opensave.ShowSave
If opensave.FileName <> "" Then
Open opensave.FileName For Output As #1
Print #1, "<html>" & Chr(13) & "<meta http-equiv=Content-Type content=text/html;charset=" & ZF & ">" & Chr(13) & "<title>" & pjtitle & "</title>" & Chr(13) & codebox.Text & Chr(13) & "</html>"
Close #1
End If
End Sub

Private Sub mlogin_Click()
If areyoulogin = False Then
   FrmLogin.Show
Else
   FrmUserMain.Show
End If
End Sub

Private Sub mNewProject_Click()
If ars = False Then
notfirhtmlpj = False
pjtitle = "新工程"
FrmMain.Caption = pjtitle & " - Firhtml2 " & firhtmlversion
ZF = "GB2312"
codebox.Text = ""
Else
saveproject.ShowSave
savefhm.ShowSave
If saveproject.FileName = "" Or savefhm.FileName = "" Then
Exit Sub
End If
pjpath = saveproject.FileName
pjmainpath = savefhm.FileName

Open pjmainpath For Output As #1
Print #1, codebox.Text
Close #1

Open pjpath For Output As #2
Print #2, pjtitle
Print #2, ZF
Print #2, pjmainpath
Close #2
notfirhtmlpj = False
pjtitle = "新工程"
FrmMain.Caption = pjtitle & " - Firhtml2 "
ZF = "GB2312"
codebox.Text = ""
End If
End Sub

Private Sub mOpenProject_Click()
Dim a, aa, S, ss, info
saveproject.ShowOpen
 If saveproject.FileName <> "" Then
   Open saveproject.FileName For Input As #1
        Do Until EOF(1)
         Line Input #1, tmpstr
         info = IIf(info = "", tmpstr, info & vbCrLf & tmpstr)
         DoEvents
        Loop
   Close #1
    a = info
    S = Split(info, vbCrLf): pjtitle = S(0): ZF = S(1): pjmainpath = S(2)
    pjpath = saveproject.FileName
    FrmMain.Caption = pjtitle & " - Firhtml2 "
   'MsgBox pjtitle & Chr(13) & ZF & Chr(13) & pjmainpath
   On Error GoTo errline
   Open pjmainpath For Input As #2
         Do Until EOF(2)
         Line Input #2, tmpstr
         codebox.Text = IIf(codebox.Text = "", tmpstr, codebox.Text & vbCrLf & tmpstr)
         DoEvents
        Loop
   Close #2
 
End If
Exit Sub

errline:
      MsgBox "工程文件未找到！", vbCritical
End Sub

Private Sub mRunInIE_Click()
  Open App.Path & "\bin.html" For Output As #1
  Print #1, "<html>" & Chr(13) & "<meta http-equiv=Content-Type content=text/html;charset=" & ZF & ">" & Chr(13) & "<title>" & pjtitle & "</title>" & Chr(13) & codebox.Text & Chr(13) & "</html>"
  Close #1
Shell "cmd /c start " & App.Path & "\bin.html"
End Sub

Private Sub mSaveOther_Click()
saveproject.ShowSave
savefhm.ShowSave
If saveproject.FileName = "" Or savefhm.FileName = "" Then
Exit Sub
End If
pjpath = saveproject.FileName
pjmainpath = savefhm.FileName

Open pjmainpath For Output As #1
Print #1, codebox.Text
Close #1

Open pjpath For Output As #2
Print #2, pjtitle
Print #2, ZF
Print #2, pjmainpath
Close #2
End Sub

Private Sub mSaveProject1_Click()
If ars = False Then
saveproject.ShowSave
savefhm.ShowSave
If saveproject.FileName = "" Or savefhm.FileName = "" Then
Exit Sub
End If
pjpath = saveproject.FileName
pjmainpath = savefhm.FileName

Open pjmainpath For Output As #1
Print #1, codebox.Text
Close #1

Open pjpath For Output As #2
Print #2, pjtitle
Print #2, ZF
Print #2, pjmainpath
Close #2
       
    ars = True
     
Else
Open pjmainpath For Output As #1
Print #1, codebox.Text
Close #1

Open pjpath For Output As #2
Print #2, pjtitle
Print #2, ZF
Print #2, pjmainpath
Close #2
     ars = True
End If
End Sub

Private Sub mYJ_Click()
FrmMyYuJu.Show
End Sub

Private Sub mZhantie_Click()
    codebox.SelText = Clipboard.GetText()
End Sub

Private Sub mZZ_Click()
FrmZZ.Show
End Sub

Private Sub projectinfo_DblClick()
If notfirhtmlpj = False Then
    projectinfo.Enabled = True
Dim a

  If projectinfo.Text = "标题" Then
     a = pjtitle
     pjtitle = InputBox("请输入工程标题")
  If pjtitle <> "" Then
     FrmMain.Caption = pjtitle & " - Firhtml2 "
  Else
     pjtitle = a
  End If

ElseIf projectinfo.Text = "字符编码" Then
     FrmZF.Show
End If
  
Else

    projectinfo.Enabled = False
    
End If

End Sub

Private Sub userkj_DblClick()
Unload FrmNewKj
If userkj.Text = "Button" Then

'加载控件新建窗口
Load FrmNewKj
FrmNewKj.Caption = "新建 Button控件"
FrmNewKj.mbotton.Visible = True
FrmNewKj.kjtype = "button"
FrmNewKj.kj_type.Text = FrmNewKj.kjtype
FrmNewKj.Visible = True
'加载完毕

ElseIf userkj.Text = "TextField" Then

'加载控件新建窗口
Load FrmNewKj
FrmNewKj.Caption = "新建 TextField控件"
FrmNewKj.mtextfield.Visible = True
FrmNewKj.kjtype = "textfield"
FrmNewKj.kj_type.Text = FrmNewKj.kjtype
FrmNewKj.Visible = True
'加载完毕

ElseIf userkj.Text = "File" Then
'加载开始
Load FrmNewKj
FrmNewKj.Caption = "新建File控件"
FrmNewKj.kjtype = "file"
FrmNewKj.mNone.Visible = True
FrmNewKj.kj_type.Text = FrmNewKj.kjtype
FrmNewKj.Visible = True
FrmNewKj.mPic.Visible = False
'加载完毕

ElseIf userkj.Text = "Label" Then
'加载开始
Load FrmNewKj
FrmNewKj.Caption = "新建Label控件"
FrmNewKj.kjtype = "label"
FrmNewKj.mNone.Visible = True
FrmNewKj.kj_type.Text = FrmNewKj.kjtype
FrmNewKj.Visible = True
FrmNewKj.mPic.Visible = False
'加载完毕

ElseIf userkj.Text = "HyperLink" Then
'加载开始
Load FrmNewKj
FrmNewKj.Caption = "新建超链接"
FrmNewKj.kjtype = "hyperlink"
FrmNewKj.kj_type.Text = FrmNewKj.kjtype
FrmNewKj.mhlink.Visible = True
FrmNewKj.Visible = True
FrmNewKj.mPic.Visible = False
'加载完毕

ElseIf userkj.Text = "Form" Then
'加载开始
Load FrmNewKj
FrmNewKj.Caption = "新建表单"
FrmNewKj.kjtype = "form"
FrmNewKj.Label1.Caption = "名称:"
FrmNewKj.Label2.Caption = "动作:"
FrmNewKj.kj_type.Locked = False
FrmNewKj.mForm.Visible = True
FrmNewKj.Visible = True
FrmNewKj.mPic.Visible = False
'加载完毕

ElseIf userkj.Text = "Picture" Then
'加载开始
Load FrmNewKj
FrmNewKj.Caption = "新建图片控件"
FrmNewKj.kjtype = "picture"
FrmNewKj.kj_type.Locked = False
FrmNewKj.Label2.Caption = "图片路径:"
FrmNewKj.mPic.Visible = True
FrmNewKj.Visible = True
'加载完毕

ElseIf userkj.Text = "Email" Then
'加载开始
Load FrmNewKj
FrmNewKj.Caption = "新建Email控件"
FrmNewKj.kjtype = "email"
FrmNewKj.Label1.Caption = "文本:"
FrmNewKj.Label2.Caption = "Email:"
FrmNewKj.kj_type.Locked = False
FrmNewKj.mNone.Visible = True
FrmNewKj.Visible = True
'加载完毕

ElseIf userkj.Text = "Br" Then
'加载开始
With codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "</br>" & Mid(.Text, .SelStart + .SelLength + 1)
End With

End If
End Sub

Private Sub wel_exit_Click()
areyoulogin = False
wel_exit.Visible = False
wel_exit2.Visible = False
End Sub

Private Sub wel_exit2_Click()
areyoulogin = False
wel_exit.Visible = False
wel_exit2.Visible = False
End Sub

Private Sub XTab1_Click()
Open App.Path & "\bin.html" For Output As #1
Print #1, codebox.Text
Close #1
codepreview.Navigate App.Path & "\bin.html"
End Sub
