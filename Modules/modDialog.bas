Attribute VB_Name = "modDialog"
Option Explicit

Public dCancel As Boolean
Public DialogWidth As Integer
Public DialogHeight As Integer
Public DialogPath As String
Public DialogView As Integer
Public dDialogFilter As String
Public DoNotUnload As Boolean

Public iCAPTION_MKDIR_TITLE As String
Public iCAPTION_MKDIR_LABLE As String
Public iCAPTION_MKDIR_OK As String
Public iCAPTION_MKDIR_CANCEL As String
Public iCAPTION_MKDIR_DEFAULTFOLDER As String

Public Const DI_NORMAL = &H3
Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)
Public Const BM_SETSTATE = &HF3
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const SRCCOPY = &HCC0020

Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Dim PrevProc As Long

Public Scrolling As Boolean
Dim Direction As Long

Public Sub Hook(hwnd As Long)
    PrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHook(hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, PrevProc
End Sub
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = 522 Then   'scroll.. wparam is a value for up or down scroll
        If wParam > 0 Then Direction = 1 Else Direction = -1
        frmDialog.FileView.MouseScroll Direction
        
    Else
    End If
        
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
End Function


