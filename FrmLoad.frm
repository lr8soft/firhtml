VERSION 5.00
Begin VB.Form FrmLoad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3225
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   360
         Top             =   2280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build 2600"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FirhtmlÍøÒ³±à¼­Æ÷"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   3105
      End
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error GoTo errline
FrmMain.Show
Timer1.Enabled = False
Unload Me
Exit Sub
errline:
  If Err.Number = "339" Then
     MsgBox "¿Ø¼þ¼ÓÔØ´íÎó£¡", vbCritical
  End If
End
End Sub
