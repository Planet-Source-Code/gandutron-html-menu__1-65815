VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   " mx2 CodeShare Explorer"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picMenuDivider 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   14280
      TabIndex        =   12
      Top             =   2040
      Width           =   14280
   End
   Begin VB.PictureBox picTitleBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   14280
      TabIndex        =   3
      Top             =   0
      Width           =   14280
      Begin VB.PictureBox picMinHover 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   11280
         Picture         =   "mdiMain.frx":038A
         ScaleHeight     =   225
         ScaleWidth      =   420
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picMaxHover 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   11640
         Picture         =   "mdiMain.frx":08B8
         ScaleHeight     =   225
         ScaleWidth      =   390
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.PictureBox picCloseHover 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   12000
         Picture         =   "mdiMain.frx":0DAA
         ScaleHeight     =   225
         ScaleWidth      =   615
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picMin 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   12840
         Picture         =   "mdiMain.frx":1530
         ScaleHeight     =   225
         ScaleWidth      =   420
         TabIndex        =   8
         Top             =   0
         Width           =   420
      End
      Begin VB.PictureBox picMax 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13200
         Picture         =   "mdiMain.frx":1A5E
         ScaleHeight     =   225
         ScaleWidth      =   390
         TabIndex        =   7
         Top             =   0
         Width           =   390
      End
      Begin VB.PictureBox picClose 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13560
         Picture         =   "mdiMain.frx":1F50
         ScaleHeight     =   225
         ScaleWidth      =   615
         TabIndex        =   6
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         Picture         =   "mdiMain.frx":26D6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblTitleBar 
         BackStyle       =   0  'Transparent
         Caption         =   "CodeShare Explorer"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   20
         Width           =   2175
      End
   End
   Begin VB.PictureBox picFrameStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1380
      ScaleWidth      =   14280
      TabIndex        =   2
      Top             =   9510
      Width           =   14280
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on a Menu Item"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   3525
      End
   End
   Begin VB.PictureBox picFrameMenu 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1790
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   14280
      TabIndex        =   0
      Top             =   255
      Width           =   14280
      Begin SHDocVwCtl.WebBrowser wbMenu 
         Height          =   1335
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   7335
         ExtentX         =   12938
         ExtentY         =   2355
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
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Win32 APIs used to toggle titlebar.
Private Declare Function GetWindowLong Lib _
   "user32" Alias "GetWindowLongA" (ByVal hwnd _
   As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib _
   "user32" Alias "SetWindowLongA" (ByVal hwnd _
   As Long, ByVal nIndex As Long, ByVal _
   dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib _
   "user32" (ByVal hwnd As Long, ByVal _
   hWndInsertAfter As Long, ByVal x As Long, _
   ByVal y As Long, ByVal cX As Long, ByVal cY _
   As Long, ByVal wFlags As Long) As Long

'Used to get window style bits.
Private Const GWL_STYLE = (-16)

'Titlebar style bit.
Private Const WS_CAPTION = &HC00000

'Force total redraw that shows new styles.
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1

Private blnResize As Boolean

Private WithEvents objMenu As HTMLDocument
Attribute objMenu.VB_VarHelpID = -1

'Variables to hold TitleBar buttons
Private btnClose As StdPicture
Private btnMax As StdPicture
Private btnMin As StdPicture

Private Function ToggleCaption(ByVal value As _
   Boolean) As Boolean
   Dim nStyle As Long

   ' Retrieve current style bits.
   nStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
   ' Set WS_SYSMENU On or Off as requested.
   If value Then
      nStyle = nStyle Or WS_CAPTION
   Else
      nStyle = nStyle And Not WS_CAPTION
   End If

   ' Try to set new style.
   If SetWindowLong(Me.hwnd, GWL_STYLE, nStyle) _
      Then
      If nStyle = GetWindowLong(Me.hwnd, _
         GWL_STYLE) Then
         ToggleCaption = True
      End If
   End If

   ' Redraw window with new style.
   SetWindowPos hwnd, 0, 0, 0, 0, 0, _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
      SWP_NOZORDER Or SWP_NOSIZE
End Function

Private Sub lblTitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If the lblTitleBar is clicked with the right mouse button, begin moving the form
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub MDIForm_Load()
    'Remove the title bar of the MDI Form
    ToggleCaption (False)
    
    'Load the default button images into variables
    Set btnClose = picClose.Picture
    Set btnMax = picMax.Picture
    Set btnMin = picMin.Picture
    
    'Loads the User interface
    'LoadUI
    SetControls
    
    'Load and process the server.xml file
    'LoadServers
    
    frmBrowser.Show
End Sub


Public Sub SetControls()

    'Set the backcolors of the forms and picture boxes
    picFrameMenu.BackColor = COLOR_APPLICATION_DARK
    
    picTitleBar.BackColor = COLOR_TITLE_BAR
    picIcon.BackColor = COLOR_TITLE_BAR
    lblTitleBar.ForeColor = vbWhite
    lblTitleBar.BackColor = COLOR_TITLE_BAR
    
    picIcon.Move picTitleBar.ScaleLeft, picTitleBar.ScaleTop
    wbMenu.Navigate App.Path & "/html/menu.html"
End Sub

Private Function objFile_oncontextmenu() As Boolean
    'disable the right click menu on the ie browser
    objFile_oncontextmenu = False
End Function

Private Function objMenu_oncontextmenu() As Boolean
    'disable the right click menu on the ie browser
    objMenu_oncontextmenu = False
End Function

Private Sub picClose_Click()
    Unload Me
End Sub

Private Sub picClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'if the mouse is over the picClose button show the hover image otherwise show the default image
    If IsOverhWnd(picClose.hwnd, x, y) Then
        picClose.Picture = picCloseHover.Picture
    Else
        'btnClose is defined on form_load
        picClose.Picture = btnClose
    End If
End Sub


Private Sub picMax_Click()
    'Set the picMax picture to the default as defined in form_load
    picMax.Picture = btnMax
    
    'If the form is maximized change it to normal, if normal then change to maximized
    If mdiMain.WindowState = 2 Then
        mdiMain.WindowState = 0
    Else
        mdiMain.WindowState = 2
    End If
End Sub

Private Sub picMax_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Same logic as picClose_MouseMove
    If IsOverhWnd(picMax.hwnd, x, y) Then
        picMax.Picture = picMaxHover.Picture
    Else
        picMax.Picture = btnMax
    End If
End Sub

Private Sub picMin_Click()
    'Minimize the form and set the image to the default image as defined in form_load
    picMin.Picture = btnMin
    mdiMain.WindowState = 1
End Sub

Private Sub picMin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Same logic as picClose_MouseMove
    If IsOverhWnd(picMin.hwnd, x, y) Then
        picMin.Picture = picMinHover.Picture
    Else
        picMin.Picture = btnMin
    End If
End Sub

Private Sub picTitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If the picTitleBar is clicked with the right mouse button, begin moving the form
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub picTitleBar_Resize()
    'Lines up the 3 control box images and set the lblTitleBar width
    picClose.Move picTitleBar.ScaleWidth - picClose.Width, picTitleBar.ScaleTop
    picMax.Move picTitleBar.ScaleWidth - picClose.Width - picMax.Width, picTitleBar.ScaleTop
    picMin.Move picTitleBar.ScaleWidth - picClose.Width - picMax.Width - picMin.Width, picTitleBar.ScaleTop

    lblTitleBar.Width = picTitleBar.ScaleWidth - picClose.Width - picMax.Width - picMin.Width - 400
End Sub

Private Sub picFrameMenu_Resize()
    'NO_ERROR_HANDLER
    On Error Resume Next
        
        'align the Browser (wbMenu) to the picFrameMenu, add +250 to move scrollbars out
        wbMenu.Move picFrameMenu.ScaleLeft, picFrameMenu.ScaleTop, picFrameMenu.ScaleWidth + 275, picFrameMenu.ScaleHeight + 275
        
    On Error GoTo 0
End Sub

Private Sub wbMenu_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim strCommand As String
    
    'If a blank URL is passed through cancel the navigation
    If URL = "http:///" Then Cancel = True
    
    'Before the browser navigates intercept the call to see if its a command
    If LCase$(Left$(URL, 8)) = "command:" Then
        'If its a command then cancel navigation
        Cancel = True
        
        'Parse the command and pass it to the menuCommandHandler
        menuCommandHandler Mid(URL, 9)
    End If
End Sub

Private Sub wbMenu_DownloadComplete()
    'Set the objMenu object to the HTML Document of the wbMenu browser
    Set objMenu = wbMenu.Document
End Sub

Private Sub menuCommandHandler(strMenuCommand As String)
    Label1.Caption = "You clicked - " & strMenuCommand
End Sub
