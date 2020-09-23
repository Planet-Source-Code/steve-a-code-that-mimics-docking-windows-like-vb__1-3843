VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7410
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   2430
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5865
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   0
      Width           =   45
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5865
      ScaleWidth      =   2430
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    'Show the two forms
    Form1.Show
    Form2.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'unload Form1 so we terminate
    Unload Form1
End Sub

Private Sub Picture1_Resize()
    'check size
    If Picture1.Width < 120 Then
        Picture1.Width = 120
    End If
    'if Form1 is docked, position it so its resizing border is hidden
    'outside the confines of Picture1.
    If Form1.bDocked Then
        Form1.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX), Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'As a simple alternative to using a splitter control we can resize
    'Picture1 by sending a message that will make windows draw a
    'resizing border for us.
    If Picture1.Visible Then
        ReleaseCapture 'need to do this or SendMessage fails
        'Send message to start resizing picture1
        SendMessage Picture1.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, ByVal &O0
    End If
End Sub
