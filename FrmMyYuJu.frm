VERSION 5.00
Begin VB.Form FrmMyYuJu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ҵ����"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4785
   Icon            =   "FrmMyYuJu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4785
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "��ͬ��"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin ����1.jcbutton b1 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "ȷ��"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.ListBox fav 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      ItemData        =   "FrmMyYuJu.frx":030A
      Left            =   240
      List            =   "FrmMyYuJu.frx":030C
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin ����1.jcbutton jcbutton1 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "ȡ��"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin ����1.jcbutton jcbutton2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "�½�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin ����1.jcbutton jcbutton3 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "ɾ��"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin ����1.jcbutton jcbutton4 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "����"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FrmMyYuJu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sts() As String
Private Sub b1_Click()
Unload Me
End Sub

Private Sub fav_Click()
With FrmMain.codebox
  .Text = Left(.Text, .SelStart + .SelLength) & fav.Text & Mid(.Text, .SelStart + .SelLength + 1)
End With
End Sub

Private Sub Form_Load()
On Error GoTo errline
    Open App.Path & "\fav.dat" For Input As #1
    i = 0
    Do Until EOF(1)
        ReDim Preserve sts(i)
        Line Input #1, sts(i)
        fav.AddItem sts(i)
        i = i + 1
    Loop
    Close #1
Exit Sub
errline:
   MsgBox "����б����ʧ�ܣ�", vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open App.Path & "\fav.dat" For Output As #1
    For i = 0 To fav.ListCount - 1
    Print #1, fav.List(i)
    Next
    Close #1
End Sub

Private Sub jcbutton1_Click()
Unload Me
End Sub

Private Sub jcbutton2_Click()
Dim a
a = InputBox("��������䣺")
If a <> "" Then
    fav.AddItem a
Else
   MsgBox "������䲻��Ϊ�գ�", vbCritical
End If
End Sub

Private Sub jcbutton3_Click()
    On Error Resume Next
    fav.RemoveItem fav.ListIndex
End Sub

Private Sub jcbutton4_Click()
MsgBox "������" & Chr(13) & "--˫�����������ʹ����䣡" & Chr(13) & "--���[�½�]�������䣡" & Chr(13) & "--���[ɾ��]��ɾ����䣡", vbInformation

End Sub
