VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4260
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
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WebBrowser.Navigate App.Path & "/html/home.html"
End Sub

Private Sub Form_Resize()
   WebBrowser.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

