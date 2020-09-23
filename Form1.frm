VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1980
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":0000
      Left            =   405
      List            =   "Form1.frx":0019
      TabIndex        =   0
      Top             =   540
      Width           =   2445
   End
   Begin Project1.FormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   503
      Caption         =   "Docking Window Example"
      RepositionForm  =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public variables used elsewhere to set values for this form's position
'and size.
Dim lFloatingWidth As Long
Dim lFloatingHeight As Long
Dim lFloatingLeft As Long
Dim lFloatingTop As Long
Dim bMoving As Boolean

'Private variables used to track moving/sizing etc.
Public bDocked As Boolean
Public lDockedWidth As Long
Public lDockedHeight As Long

Private Sub Form_Load()
    'Initialize the positions/sizes of this form
    lDockedWidth = MDIForm1.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = MDIForm1.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    lFloatingLeft = Me.Left
    lFloatingTop = Me.Top
    lFloatingWidth = Me.Width
    lFloatingHeight = Me.Height
    'Start with the form docked in Picture1 on the MDI Form
    'put Form1 in the 'Dock' and position it so its resizing border is
    'hidden outside the confines of Picture1
    bDocked = True
    SetParent Me.hwnd, MDIForm1!Picture1.hwnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    MDIForm1!Picture1.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'reset this form's owner to prevent a crash
    Call SetWindowWord(Me.hwnd, SWW_HPARENT, 0&)
End Sub

Private Sub Form_Resize()

    If Me.WindowState <> vbMinimized Then
        'Update the stored Values
        StoreFormDimensions
        'position and size the listbox
        List1.Move 3 * Screen.TwipsPerPixelX, FormDragger1.Height + (3 * Screen.TwipsPerPixelY), Me.ScaleWidth - (7 * Screen.TwipsPerPixelX), Me.ScaleHeight - (FormDragger1.Height + (6 * Screen.TwipsPerPixelY))
    End If
    
End Sub

Private Sub FormDragger1_DblClick()

    'Snap the form in or out of the dock (Picture1)
    bMoving = True 'stop the new dimensions being stored
    If bDocked Then
        'Undock
        Me.Visible = False
        bDocked = False
        SetParent Me.hwnd, 0
        Me.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
        MDIForm1!Picture1.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hwnd, SWW_HPARENT, MDIForm1.hwnd)
    Else
        'Dock
        bDocked = True
        SetParent Me.hwnd, MDIForm1!Picture1.hwnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        MDIForm1!Picture1.Visible = True
    End If
    bMoving = False

End Sub

Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    
    Dim rct As RECT

    'If over Picture1 on MDIForm1 which we are using as a Dock, set parent
    'of this form to Picture1, and position it at -4,-4 pixels, otherwise
    'set this Form's parent to the desktop and postion it at Left,Top
    'We dont need to size the form, as the DragForm control will have done
    'this for us.
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    'Get the screen based coordinates of Picture1
    GetWindowRect MDIForm1!Picture1.hwnd, rct
    'Inflate the rect because we want the form to be bigger than Picture1
    'to hide it's border
    With rct
        .Left = .Left - 4
        .Top = .Top - 4
        .Right = .Right + 4
        .Bottom = .Bottom + 4
    End With
    'See if the top/left corner of this form is in Picture1's screen rectangle
    'As we have set RepositionForm to false, we are responsible for positioning the form
    If PtInRect(rct, FormLeft, FormTop) Then
        bDocked = True
        SetParent Me.hwnd, MDIForm1!Picture1.hwnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        MDIForm1!Picture1.Visible = True
    Else
        Me.Visible = False
        bDocked = False
        SetParent Me.hwnd, 0
        Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
        MDIForm1!Picture1.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hwnd, SWW_HPARENT, MDIForm1.hwnd)
    End If
    
    'reset the moving flag and store the form dimensions
    bMoving = False
    StoreFormDimensions

End Sub

Private Sub FormDragger1_FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    
    Dim rct As RECT
    
    'Set the moving flag so we dont store the wrong dimensions
    bMoving = True
    
    'If over Picture1 on MDIForm1 which we are using as a Dock, change the width to that of
    'Picture1, else change it to the 'floating width and height
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    'Get the screen based coordinates of Picture1
    GetWindowRect MDIForm1!Picture1.hwnd, rct
    'Inflate the rect because we want the form to be bigger than Picture1
    'to hide it's border
    With rct
        .Left = .Left - 4
        .Top = .Top - 4
        .Right = .Right + 4
        .Bottom = .Bottom + 4
    End With
    'See if the top/left corner of this form is in Picture1's screen rectangle
    
    If PtInRect(rct, FormLeft, FormTop) Then
        FormWidth = lDockedWidth / Screen.TwipsPerPixelX
        FormHeight = lDockedHeight / Screen.TwipsPerPixelY
    Else
        FormWidth = lFloatingWidth / Screen.TwipsPerPixelX
        FormHeight = lFloatingHeight / Screen.TwipsPerPixelY
    End If

End Sub

Private Sub StoreFormDimensions()

   'Store the height/width values
    If Not bMoving Then
        If bDocked Then
            lDockedWidth = Me.Width
            lDockedHeight = Me.Height
        Else
            lFloatingLeft = Me.Left
            lFloatingTop = Me.Top
            lFloatingWidth = Me.Width
            lFloatingHeight = Me.Height
        End If
    End If
End Sub

