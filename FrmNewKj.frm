VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmNewKj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÐÂ½¨ Kj_Name..."
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   Icon            =   "FrmNewKj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6750
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox mPic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   4905
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox picchang 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox pickuan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   2175
      End
      Begin ¹¤³Ì1.jcbutton jcbutton3 
         Height          =   855
         Left            =   3480
         TabIndex        =   31
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Í¼Æ¬Â·¾¶"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Í¼Æ¬¿í:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Í¼Æ¬³¤:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   120
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog openpic 
      Left            =   5640
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ËùÓÐÎÄ¼þ(*.*)|*.*|"
   End
   Begin VB.PictureBox mForm 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   4905
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton Option8 
         Caption         =   "GET"
         Height          =   180
         Left            =   1560
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "POST"
         Height          =   180
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "±íµ¥·½·¨:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.PictureBox mhlink 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   4905
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin RichTextLib.RichTextBox linktext 
         Height          =   855
         Left            =   720
         TabIndex        =   21
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1508
         _Version        =   393217
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"FrmNewKj.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÍøÖ·:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.PictureBox mNone 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   4905
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÎÞ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   720
         TabIndex        =   18
         Top             =   120
         Width           =   210
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÉèÖÃ:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.PictureBox mtextfield 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   4905
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton Option6 
         Caption         =   "µ¥ÐÐ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   14
         Top             =   170
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "ÃÜÂë"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "¶àÐÐ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   12
         Top             =   170
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÀàÐÍ:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.PictureBox mbotton 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   4905
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton Option3 
         Caption         =   "ÎÞ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   10
         Top             =   170
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "ÖØÖÃ±íµ¥"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ìá½»±íµ¥"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Top             =   170
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¶¯×÷:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   465
      End
   End
   Begin ¹¤³Ì1.jcbutton jcbutton1 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "È·¶¨"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.TextBox kj_type 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox kj_text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin ¹¤³Ì1.jcbutton jcbutton2 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "È¡Ïû"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Ø¼þÃû³Æ:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Öµ(Value):"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FrmNewKj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kjtype
Dim choice



Private Sub Form_Unload(Cancel As Integer)
kjtype = ""
End Sub

Private Sub jcbutton1_Click()
   Dim btype, bname, bValue, bmethod, bdongzuo, bpath
If kjtype = "button" Then
  If choice <> "" Then
     If choice = 1 Then btype = "submit"
     If choice = 2 Then btype = "reset"
     If choice = 3 Then btype = "button"
    ' MsgBox btype
   

     bname = kj_text.Text
     bValue = kj_type.Text
           With FrmMain.codebox
           .Text = Left(.Text, .SelStart + .SelLength) & "<input type=" & """" & btype & """" & " name=" & """" & "Submit" & """" & " value=" & """" & bname & """" & " />" & Mid(.Text, .SelStart + .SelLength + 1)
           End With
   
  End If
  
ElseIf kjtype = "textfield" Then
     bname = kj_text.Text
     bValue = kj_type.Text
  If choice <> "" Then
     If choice = 4 Then btype = "text"
     '¿ªÊ¼
     If choice = 5 Then
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<textarea name=" & """" & bname & """" & " id=" & """" & bname & """" & "></textarea>" & Mid(.Text, .SelStart + .SelLength + 1)
         End With
     Unload Me
     Exit Sub
     End If
     '½áÊø
     If choice = 6 Then btype = "password"
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<input name=" & """" & bValue & """" & " type=" & """" & btype & """" & " value=" & """" & bname & """" & " id=" & bValue & " />" & Mid(.Text, .SelStart + .SelLength + 1)
         End With
     End If
  
  ElseIf kjtype = "file" Then
     bname = kj_text.Text
     bValue = kj_type.Text
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<input type=" & """" & bValue & """" & " name=" & """" & bname & """" & "  />" & Mid(.Text, .SelStart + .SelLength + 1)
         End With


  ElseIf kjtype = "label" Then
     bname = kj_text.Text
     bValue = kj_type.Text
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<label>" & bname & "</label>" & Mid(.Text, .SelStart + .SelLength + 1)
        End With
  
  ElseIf kjtype = "hyperlink" Then
     bname = kj_text.Text
     bValue = kj_type.Text
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<a href=" & """" & linktext.Text & """" & ">" & bname & "</a>" & Mid(.Text, .SelStart + .SelLength + 1)
        End With
        
  ElseIf kjtype = "form" Then
    If Option7 = True Or Option8 = True Then
         If Option7 = True Then bmethod = "POST"
         If Option8 = True Then bmethod = "GET"
    Else
         bmethod = "POST"
    End If
     bname = kj_text.Text
     bdongzuo = kj_type.Text
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<form id=" & """" & bname & """" & " name=" & """" & bname & """" & " method=" & """" & bmethod & """" & " action=" & """" & bdongzuo & """" & ">" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "</form>" & Mid(.Text, .SelStart + .SelLength + 1)
        End With
        
    
  ElseIf kjtype = "picture" Then
     bname = kj_text.Text
     bpath = kj_type.Text
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<input name=" & """" & bname & """" & " type=" & """" & "image" & """" & " src=" & """" & bpath & """" & " />" & Mid(.Text, .SelStart + .SelLength + 1)
        End With
         
  ElseIf kjtype = "email" Then
    Dim btext, bemail
     btext = kj_text.Text
     bemail = kj_type.Text
         With FrmMain.codebox
         .Text = Left(.Text, .SelStart + .SelLength) & "<a href=" & """" & "mailto:" & bemail & """" & ">" & btext & "</a>" & Mid(.Text, .SelStart + .SelLength + 1)
         End With
         
  End If
'±£´æÁÙÊ±ÎÄ¼þ
Open App.Path & "\bin.html" For Output As #1
Print #1, FrmMain.codebox.Text
Close #1
FrmMain.codepreview.Navigate App.Path & "\bin.html"
Unload Me
End Sub

Private Sub jcbutton2_Click()
Unload Me

End Sub

Private Sub jcbutton3_Click()
openpic.ShowOpen
kj_type.Text = openpic.FileName
End Sub

Private Sub Option2_Click()
choice = 2
End Sub

Private Sub Option3_Click()
choice = 3
End Sub

Private Sub Option1_Click()
choice = 1
End Sub

Private Sub Option4_Click()
choice = 5
End Sub

Private Sub Option5_Click()
choice = 6
End Sub

Private Sub Option6_Click()
choice = 4
End Sub
