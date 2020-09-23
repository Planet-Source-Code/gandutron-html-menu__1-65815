Attribute VB_Name = "modCodeShare"
Option Explicit

Public Const COLOR_APPLICATION_DARK = 11770523
Public Const COLOR_APPLICATION_LIGHT = 13285552
Public Const COLOR_TITLE_BAR = 3871790

Public codeShareURL As String
Public userId As String

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Size/Move Constants
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const HTBOTTOMRIGHT = 17
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Function from http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=62187&lngWId=1
Public Function IsOverhWnd(hwnd As Long, x As Single, y As Single) As Boolean
Dim Rec As RECT

    'get control position within the desktop
    If GetWindowRect(hwnd, Rec) = 0 Then Exit Function
    
    'x & y are currently in twips, so convert them to pixels
    x = x / Screen.TwipsPerPixelX
    y = y / Screen.TwipsPerPixelY
    
    'check if cursor is over the control
    If (x < 0) Or (y < 0) Or (x > Rec.Right - Rec.Left) Or (y > Rec.Bottom - Rec.Top) Then
        ReleaseCapture 'stop capturing the mouse
        IsOverhWnd = False
       Else
        SetCapture hwnd 'capture the mouse leaving the control
        IsOverhWnd = True
    End If
    
End Function

'*****************************************************************************
'* ZeroIfNegative
'*****************************************************************************
Public Function ZeroIfNegative(ByVal value As Long) As Long
   ' Returns Zero if the value is negative mainly used by usercontrol_resize events
   If value > 0 Then
      ZeroIfNegative = value
   Else
      ZeroIfNegative = 0
   End If
   'alternatively you could write: ZeroIfNegative = IIf(Value > 0, Value, 0)
   'but the "old" if statement is much faster
End Function
