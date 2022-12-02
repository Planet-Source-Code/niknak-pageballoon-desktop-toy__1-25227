VERSION 5.00
Begin VB.Form frm_balloon 
   BorderStyle     =   0  'None
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tim_move 
      Interval        =   64
      Left            =   360
      Top             =   420
   End
   Begin VB.Image img_shapemap 
      Height          =   2130
      Index           =   0
      Left            =   1680
      Picture         =   "frm_balloon.frx":0000
      Top             =   0
      Width           =   840
   End
   Begin VB.Image img_shapemap 
      Height          =   2130
      Index           =   1
      Left            =   840
      Picture         =   "frm_balloon.frx":5D74
      Top             =   0
      Width           =   840
   End
   Begin VB.Image img_shapemap 
      Height          =   2130
      Index           =   2
      Left            =   0
      Picture         =   "frm_balloon.frx":BAE8
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "frm_balloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************
'API DECLARATIONS
'************************************************************
'USED TO KEEP FORM ONTOP
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
'*****************************************
'PRIVATE VARIABLES
'*****************************************
Dim caught As Boolean
Dim ypos As Long
Dim xpos As Long

'*****************************************
'FORM LOAD ROUTINE
'*****************************************
Private Sub Form_Load()
    Dim rndmap As Integer

    busy = True
    frm_wait.Show
    DoEvents
    
    Randomize
    rndmap = Int((3) * Rnd)
    
    Select Case rndmap
        Case Is = 0
            frm_wait.lbl_wait.Caption = "Creating Page Balloon..."
        Case Is = 1
            frm_wait.lbl_wait.Caption = "Creating Moon Balloon..."
        Case Is = 2
            frm_wait.lbl_wait.Caption = "Creating Frank Balloon..."
    End Select
    
    'SKIN THIS FORM
    If Not verify_file(App.Path & "\" & Me.Name & Str(rndmap) & ".tmp") Then
        SavePicture img_shapemap(rndmap).Picture, App.Path & "\" & Me.Name & Str(rndmap) & ".tmp"
    End If
    
    face = CreateRegionFromFile(Me, img_shapemap(rndmap), App.Path & "\" & Me.Name & Str(rndmap) & ".tmp", RGB(0, 255, 0))
    SetWindowRgn Me.hwnd, face, True
    
    img_shapemap(0).Visible = False
    img_shapemap(1).Visible = False
    img_shapemap(2).Visible = False

    Me.Show
    xpos = Me.Left
    ypos = Me.Top
    Unload frm_wait
    busy = False
End Sub

'*****************************************
'FORM MOVING AND POPPING ROUTINE
'*****************************************
Private Sub form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        caught = True
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
        caught = False
    Else
        Beep 600, 30
        Sleep 100
        Unload Me
    End If
End Sub

'*****************************************
'MOVE TIMER MOVES THE BALLOON AT SET INTERVALS
'*****************************************
Public Sub tim_move_Timer()
    If Not busy Then
        If frm_main.chk_wrap Then
            If ypos > -Me.Height Then
                ypos = Me.Top - 20
            Else
                ypos = Screen.Height
            End If
        Else
            If ypos > 0 Then
                ypos = Me.Top - 20
            Else
                ypos = 0
            End If
        End If
        If Not caught Then
            If frm_main.chk_ontop Then
                SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, ypos / Screen.TwipsPerPixelY, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE
            Else
                SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.Left / Screen.TwipsPerPixelX, ypos / Screen.TwipsPerPixelY, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE
            End If
        End If
    End If
End Sub
