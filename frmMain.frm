VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNavigate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "http://www.planet-source-code.com/vb/"
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Navigate "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8895
      ExtentX         =   15690
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNavigate_Click()
WebBrowser.Navigate txtNavigate
End Sub

Private Sub Form_Load()
    'Start Trapping Right-Mouse clicks in WebBrowser Control:
    gLngMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, App.hInstance, GetCurrentThreadId)

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Cancel the trapping of the code
UnhookWindowsHookEx gLngMouseHook

End Sub


