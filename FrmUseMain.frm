VERSION 5.00
Begin VB.Form FrmUserMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”√ªßπ‹¿Ì"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Œ¢»Ì—≈∫⁄"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUseMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4995
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "%username%"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ª∂”≠ªÿ¿¥"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1140
   End
End
Attribute VB_Name = "FrmUserMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = username
End Sub

