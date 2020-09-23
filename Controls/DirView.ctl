VERSION 5.00
Begin VB.UserControl DirView 
   BackColor       =   &H005A371B&
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   Begin VB.PictureBox DropDownIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   1680
      Picture         =   "DirView.ctx":0000
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2760
      Top             =   480
   End
   Begin VB.PictureBox PicDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      Begin OricronDialog.Button picBtn 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
      End
   End
   Begin VB.PictureBox PicContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.PictureBox picOzadje 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   360
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   241
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   11
         Top             =   0
         Width           =   3015
         Begin VB.PictureBox picSel 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            FillColor       =   &H003E2911&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   12
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox IconBuff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox List5 
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picBuffText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "DirView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:Oricron (Almoust all, all the ui code is mine, just took a peak into regestry and geticon code;)

'Modify as you want, if you have time, vote
'or at least coment...
'if you built it into your app, tell me to
'so i can take a look;)

Option Explicit
'Local Variables
Dim Spacing As Integer      'space left for the folder from it's parent form
Dim iIcon As Long           'icon of the folder we get from shell
Dim iSelected As Integer    'selected row
Dim tSel As Integer
Dim InFocus As Boolean      'we will need this to check when we draw to display picturebox

'FUNCTIONS:
'Graphics
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

'
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80

'To get icons from files...
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long

Public Path As String
Public DesktopFolder As String
Public MyDocumentsFolder As String
Public MyComputer As String
Public MyDocuments As String

'Colors
Public SelBorderColor As OLE_COLOR
Public SelBackColor As OLE_COLOR
Public BorderColor As OLE_COLOR

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'regestriy stuff
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4
Const ERROR_SUCCESS = 0&

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'eVENTS
Public Event Click(index As Integer, Path As String)
Public Event ErrorOcured(Error As Long)

'This code is written by Tim Misset - to get icons from folders i took some code from him - thanks;)
'Const LARGE_ICON As Integer = 32                '  do not need that so i removed it
'Const SMALL_ICON As Integer = 16                '  do not need that so i removed it
Const MAX_PATH = 260
Const ILD_TRANSPARENT = &H1                      '  Display transparent
Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Const SHGFI_EXETYPE = &H2000                     '  return exe type
'Const SHGFI_LARGEICON = &H0                     '  get large icon - do not need that so i removed it
Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Const SHGFI_SMALLICON = &H1                      '  get small icon
Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Const SHGFI_TYPENAME = &H400                     '  get type name
Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE


Private Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long                      '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80          '  out: type name
End Type
 
Dim SHInfo As SHFILEINFO


'AlphaBlending - to get blendet colors (buttons etc.)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)

Private Type ColorAndAlpha
    r                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
End Type

Private Function AlphaBlend(ByVal FirstColor As Long, ByVal SecondColor As Long, ByVal AlphaValue As Long) As Long
    Dim iForeColor         As ColorAndAlpha
    Dim iBackColor         As ColorAndAlpha
    
    OleTranslateColor FirstColor, 0, VarPtr(iForeColor)
    OleTranslateColor SecondColor, 0, VarPtr(iBackColor)
    With iForeColor
        .r = (.r * AlphaValue + iBackColor.r * (255 - AlphaValue)) / 255
        .G = (.G * AlphaValue + iBackColor.G * (255 - AlphaValue)) / 255
        .B = (.B * AlphaValue + iBackColor.B * (255 - AlphaValue)) / 255
    End With
    CopyMemory VarPtr(AlphaBlend), VarPtr(iForeColor), 4
    
End Function
Private Sub Generate()
Dim ICNT As Integer
Dim IcNT2 As Integer

Dim caption As String
Dim iPos As Integer
Dim X1 As Integer
Dim ipath As String
Dim exists As Boolean

exists = CheckPath

'If Right(Path, 1) = "\" Then
'    Path = left(Path, Len(Path) - 1)
'End If

ipath = Path

Dim isel As Integer
isel = -1
Dim RenderFolders As Boolean
RenderFolders = True
picOzadje.Width = 1
'Here we generate what the user sees
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear

'create dektop "folder"
Dim Desktop, DesktopPath As String

DesktopPath = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
Desktop = Mid(DesktopPath, InStrRev(DesktopPath, "\") + 1)

DesktopFolder = DesktopPath
List1.AddItem Desktop
List2.AddItem 0
List3.AddItem DesktopPath

MyDocuments = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}", "")

If MyDocuments = "" Then MyDocuments = "My Documents"

MyDocumentsFolder = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")

List1.AddItem MyDocuments
List2.AddItem 1
List3.AddItem MyDocumentsFolder

Dim RenderDesktop As Boolean
RenderDesktop = False

If Capitalize(Left(Path, Len(DesktopPath))) = Capitalize(DesktopPath) Then
    RenderDesktop = True
    RenderFolders = False
End If

If Capitalize(Left(Path, Len(MyDocumentsFolder))) = Capitalize(MyDocumentsFolder) And exists = True Then
    Do Until Len(ipath) <= Len(MyDocumentsFolder)
        iPos = InStrRev(ipath, "\")
        List4.AddItem Mid(ipath, iPos + 1), 0
        List5.AddItem ipath, 0
        ipath = Left(ipath, iPos - 1)
            
    Loop
        
    For IcNT2 = 0 To List4.ListCount - 1
        List1.AddItem List4.List(IcNT2)
        List2.AddItem 2 + IcNT2
        List3.AddItem List5.List(IcNT2)
    Next IcNT2
    
    isel = List1.ListCount - 1

    RenderFolders = False
End If

'create mycomputer "folder"

MyComputer = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "")
If MyComputer = "" Then MyComputer = "My Computer"

List1.AddItem MyComputer
List2.AddItem 1
List3.AddItem "#MYCOMPUTER"

'add the remaining folers to the view


For ICNT = 0 To Drive1.ListCount - 1
    caption = Mid(Drive1.List(ICNT), 5)
    If caption <> "" Then
        caption = Left(caption, Len(caption) - 1) & " "
    
    End If
    
    List1.AddItem caption & "(" & Capitalize(Left(Drive1.List(ICNT), 1)) & ":)"
  
    List2.AddItem 2
    List3.AddItem Capitalize(Left(Drive1.List(ICNT), 1)) & ":\"
  
    iPos = 4
    X1 = List3.ListCount - 1
    If List3.List(X1) = Capitalize(Left(Path, 3)) And RenderFolders = True And exists = True And Len(Path) > 3 Then
        Do Until iPos < 4
            iPos = InStrRev(ipath, "\")
            List4.AddItem Mid(ipath, iPos + 1), 0
            List5.AddItem ipath, 0
            ipath = Left(ipath, iPos - 1)
            
        Loop
        
        For IcNT2 = 0 To List4.ListCount - 1
            List1.AddItem List4.List(IcNT2)
            List2.AddItem 3 + IcNT2
            List3.AddItem List5.List(IcNT2)
        Next IcNT2
        isel = List1.ListCount - 1
        
    End If
Next ICNT

'To get folders on the desktop
Dir1.Path = DesktopPath

For ICNT = 0 To Dir1.ListCount - 1
    List1.AddItem Mid(Dir1.List(ICNT), InStrRev(Dir1.List(ICNT), "\") + 1)
    List2.AddItem 1
    List3.AddItem Dir1.List(ICNT)
    
    If RenderDesktop = True And exists = True Then
        If Capitalize(Left(Path, Len(Dir1.List(ICNT)) + 1)) = Capitalize(Dir1.List(ICNT) & "\") Then
            
            Do Until Len(ipath) <= Len((Dir1.List(ICNT)))
                iPos = InStrRev(ipath, "\")
                List4.AddItem Mid(ipath, iPos + 1), 0
                List5.AddItem ipath, 0
                ipath = Left(ipath, iPos - 1)
                    
            Loop
                
            For IcNT2 = 0 To List4.ListCount - 1
                List1.AddItem List4.List(IcNT2)
                List2.AddItem 3 + IcNT2
                List3.AddItem List5.List(IcNT2)
            Next IcNT2
            
            isel = List1.ListCount - 1
        End If
        
    End If
Next ICNT

'now we have the propertys and we print them to the picturebox
picBuffText.Height = picBuffText.TextHeight("Žg")
picBuff.Height = picBuffText.Height + 2

picOzadje.Height = List1.ListCount * picBuff.Height + 1
picOzadje.Cls
picBuffText.BackColor = vbWhite
picBuffText.ForeColor = vbBlack

For ICNT = 0 To List1.ListCount - 1
    'here we print the caption of the line
    picBuffText.Cls
    picBuffText.Width = picBuffText.TextWidth(List1.List(ICNT))
    picBuffText.Print List1.List(ICNT)

    
    picBuff.Width = Spacing * List2.List(ICNT) + picBuffText.Width + 18 + 3
    
    'here we copy the text to the buffer picture (one line in the view)
    picBuff.Cls
    BitBlt picBuff.hDC, Spacing * List2.List(ICNT) + 18 + 2, 1, picBuffText.Width, picBuff.Height, picBuffText.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    'If the background picture is not wide enough we resize it to match our needs
    If picOzadje.Width < picBuff.Width + 1 Then picOzadje.Width = picBuff.Width + 1
    
    'now we copy curent line to the list
    BitBlt picOzadje.hDC, 0, picBuff.Height * ICNT, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
    
    'to add an icon infront of the caption
    IconBuff.Cls
    
Dim hLargeIcon As Long
Dim hSmallIcon As Long
                
    If List3.List(ICNT) = "#MYCOMPUTER" Then
        ExtractIconEx "explorer.exe", 0, hLargeIcon, hSmallIcon, 1
        DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
        DestroyIcon hSmallIcon
        DestroyIcon hLargeIcon
    ElseIf List3.List(ICNT) = MyDocumentsFolder Then
        ExtractIconEx "mydocs.dll", 0, hLargeIcon, hSmallIcon, 1
        DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
        DestroyIcon hSmallIcon
        DestroyIcon hLargeIcon
    ElseIf List3.List(ICNT) = DesktopFolder Then
        ExtractIconEx "shell32.dll", 34, hLargeIcon, hSmallIcon, 1
        DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
        DestroyIcon hSmallIcon
        DestroyIcon hLargeIcon
    Else
        iIcon = SHGetFileInfo(List3.List(ICNT), 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
    End If
    
    BitBlt picOzadje.hDC, Spacing * List2.List(ICNT) + 2, picBuff.Height * ICNT, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY

Next ICNT
    
If isel < 0 Then
    For ICNT = 0 To List3.ListCount - 1
        If List3.List(ICNT) = Path Then isel = ICNT
    Next ICNT
End If

If picOzadje.Width < UserControl.Width / Screen.TwipsPerPixelX - 2 Then picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
PicContainer.Width = (picOzadje.Width + 2) * Screen.TwipsPerPixelX
PicContainer.Height = (picOzadje.Height + 2) * Screen.TwipsPerPixelY
SelectIndex isel
tSel = isel

End Sub

Public Function GetValue(ByVal hKey As Long, _
ByVal strPath As String, ByVal strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

If Not IsEmpty(Default) Then
GetValue = Default
Else
GetValue = ""
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetValue = Left$(strBuffer, intZeroPos - 1)
Else
GetValue = strBuffer
End If

End If

Else
' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function


Public Sub Refresh()
Generate

End Sub

Private Sub picBtn_Click()
Dim Rec As RECT

GetWindowRect UserControl.hwnd, Rec

PicContainer.Top = (Rec.Bottom) * Screen.TwipsPerPixelY
PicContainer.Left = Rec.Left * Screen.TwipsPerPixelX
Generate


PicContainer.ZOrder
SelectIndex tSel
PicContainer.Visible = True

SetCapture picOzadje.hwnd

End Sub

Private Sub PicDisplay_Click()
picBtn_Click

End Sub

Private Sub PicDisplay_Resize()
picBtn.Left = UserControl.Width / Screen.TwipsPerPixelX - picBtn.Width - 1
picBtn.Top = (PicDisplay.Height - picBtn.Height) / 2

PicDisplay.BackColor = vbWhite
PicDisplay.ForeColor = BorderColor
PicDisplay.Line (0, 0)-(PicDisplay.Width, 0)
PicDisplay.Line (0, 0)-(0, PicDisplay.Height - 1)
PicDisplay.Line (0, PicDisplay.Height - 1)-(PicDisplay.Width, PicDisplay.Height - 1)
PicDisplay.Line (PicDisplay.Width - 1, 0)-(PicDisplay.Width - 1, PicDisplay.Height - 1)

picBuff.Cls
On Error Resume Next
picBuff.Width = picBtn.Left - 2
BitBlt picBuff.hDC, 0, (picBuff.Height - IconBuff.Height) / 2, picBuff.Width, picBuff.Height, picOzadje.hDC, (List2.List(tSel) * Spacing - 1), tSel * picBuff.Height, SRCCOPY

'now copy the whole thing to the displayed picturebox
BitBlt PicDisplay.hDC, 2, (PicDisplay.Height - picBuff.Height) / 2, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY

PicDisplay.Refresh

End Sub

Private Sub picOzadje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Rec As RECT, ok, lx, ly

ok = GetClientRect(PicContainer.hwnd, Rec)

lx = CLng(X)
ly = CLng(Y)

If PtInRect(Rec, lx, ly) = 0 Then  'Returns 1 if true, 0 if false
    PicContainer.Visible = False
    ReleaseCapture
    Generate ' we pass this so that the shown path is the path of the dirview
Else
    PicContainer.Visible = False
    ReleaseCapture
    
    picBuff.Cls
    picBuff.Width = picBtn.Left - 2
    
    BitBlt picBuff.hDC, 0, (picBuff.Height - IconBuff.Height) / 2, picBuff.Width, picBuff.Height, picSel.hDC, (List2.List(iSelected) * Spacing - 1), 0, SRCCOPY
    
    'now copy the whole thing to the displayed picturebox
    BitBlt PicDisplay.hDC, 2, (PicDisplay.Height - picBuff.Height) / 2, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
    PicDisplay.Refresh

    RaiseEvent Click(iSelected, List3.List(iSelected))
End If

End Sub

Private Sub picOzadje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X < picOzadje.Width And Y >= 0 And Y < picOzadje.Height Then
    If iSelected <> Int(Y / picBuff.Height) Then
        SelectIndex Int(Y / picBuff.Height)
    End If
End If
End Sub

Private Sub picOzadje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = vbLeftButton Then RaiseEvent Click(iSelected, List3.List(iSelected))

End Sub

Private Sub UserControl_EnterFocus()
picBuff.Cls
picBuff.Width = picBtn.Left - 2

BitBlt picBuff.hDC, 0, (picBuff.Height - IconBuff.Height) / 2, picBuff.Width, picBuff.Height, picSel.hDC, (List2.List(tSel) * Spacing - 1), 0, SRCCOPY

'now copy the whole thing to the displayed picturebox
BitBlt PicDisplay.hDC, 2, (PicDisplay.Height - picBuff.Height) / 2, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
PicDisplay.Refresh

InFocus = True

End Sub

Private Sub UserControl_ExitFocus()
picBuff.Cls
picBuff.Width = picBtn.Left - 2

BitBlt picBuff.hDC, 0, (picBuff.Height - IconBuff.Height) / 2, picBuff.Width, picBuff.Height, picOzadje.hDC, (List2.List(tSel) * Spacing - 1), tSel * picBuff.Height, SRCCOPY

'now copy the whole thing to the displayed picturebox
BitBlt PicDisplay.hDC, 2, (PicDisplay.Height - picBuff.Height) / 2, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
PicDisplay.Refresh

InFocus = False

End Sub

Public Sub UserControl_Initialize()
SelBorderColor = vbHighlight
SelBackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)

Spacing = 10
picSel.Left = 0
PicDisplay.Top = 0
PicDisplay.Left = 0

picBuff.Height = PicDisplay.Height - 2
picBuff.Width = 13

picBuff.BackColor = &H8000000F

picBuff.ForeColor = &H80000014
picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
picBuff.Line (0, 0)-(0, picBuff.Height - 1)

picBuff.ForeColor = &H80000010
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

DropDownIcon.BackColor = picBuff.BackColor

BitBlt picBuff.hDC, (picBuff.Width - DropDownIcon.Width) / 2, (picBuff.Height - DropDownIcon.Height) / 2, DropDownIcon.Width, DropDownIcon.hwnd, DropDownIcon.hDC, 0, 0, SRCCOPY

Set picBtn.NormalImage = picBuff.Image
Set picBtn.FocusedImage = picBuff.Image

picBuff.Cls
picBuff.BackColor = &H8000000F

picBuff.ForeColor = &H80000010
picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
picBuff.Line (0, 0)-(0, picBuff.Height - 1)
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

DropDownIcon.BackColor = picBuff.BackColor

BitBlt picBuff.hDC, (picBuff.Width - DropDownIcon.Width) / 2 + 1, (picBuff.Height - DropDownIcon.Height) / 2 + 1, DropDownIcon.Width, DropDownIcon.hwnd, DropDownIcon.hDC, 0, 0, SRCCOPY

Set picBtn.PressedImage = picBuff.Image

picBuff.BackColor = vbWhite

SetWindowLong PicContainer.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
SetParent PicContainer.hwnd, 0

BorderColor = AlphaBlend(AlphaBlend(vbButtonFace, vbWhite, 170), vb3DDKShadow, 70)

Path = "C:\"
Generate

picOzadje.Top = 1
picOzadje.Left = 1

End Sub

Private Function Capitalize(Text As String) As String
Dim ICNT As Integer

'useful sometimes - if the path is not send as it
'really is (caps) example you say c:\wIndows instead of C:\WINDOWS
'CHECKS ONLY THE ENGLISH ALPHABETH, AND SOME MORE (SLO, CRO, GER)...

For ICNT = 1 To Len(Text)
    If Mid(Text, ICNT, 1) = "a" Then
        Capitalize = Capitalize & "A"
    ElseIf Mid(Text, ICNT, 1) = "b" Then
        Capitalize = Capitalize & "B"
    ElseIf Mid(Text, ICNT, 1) = "c" Then
        Capitalize = Capitalize & "C"
    ElseIf Mid(Text, ICNT, 1) = "è" Then
        Capitalize = Capitalize & "È"
    ElseIf Mid(Text, ICNT, 1) = "æ" Then
        Capitalize = Capitalize & "Æ"
    ElseIf Mid(Text, ICNT, 1) = "d" Then
        Capitalize = Capitalize & "D"
    ElseIf Mid(Text, ICNT, 1) = "e" Then
        Capitalize = Capitalize & "E"
    ElseIf Mid(Text, ICNT, 1) = "f" Then
        Capitalize = Capitalize & "F"
    ElseIf Mid(Text, ICNT, 1) = "g" Then
        Capitalize = Capitalize & "G"
    ElseIf Mid(Text, ICNT, 1) = "h" Then
        Capitalize = Capitalize & "H"
    ElseIf Mid(Text, ICNT, 1) = "i" Then
        Capitalize = Capitalize & "I"
    ElseIf Mid(Text, ICNT, 1) = "j" Then
        Capitalize = Capitalize & "J"
    ElseIf Mid(Text, ICNT, 1) = "k" Then
        Capitalize = Capitalize & "K"
    ElseIf Mid(Text, ICNT, 1) = "l" Then
        Capitalize = Capitalize & "L"
    ElseIf Mid(Text, ICNT, 1) = "m" Then
        Capitalize = Capitalize & "M"
    ElseIf Mid(Text, ICNT, 1) = "n" Then
        Capitalize = Capitalize & "N"
    ElseIf Mid(Text, ICNT, 1) = "o" Then
        Capitalize = Capitalize & "O"
    ElseIf Mid(Text, ICNT, 1) = "p" Then
        Capitalize = Capitalize & "P"
    ElseIf Mid(Text, ICNT, 1) = "r" Then
        Capitalize = Capitalize & "R"
    ElseIf Mid(Text, ICNT, 1) = "q" Then
        Capitalize = Capitalize & "Q"
    ElseIf Mid(Text, ICNT, 1) = "s" Then
        Capitalize = Capitalize & "S"
    ElseIf Mid(Text, ICNT, 1) = "š" Then
        Capitalize = Capitalize & "Š"
    ElseIf Mid(Text, ICNT, 1) = "t" Then
        Capitalize = Capitalize & "T"
    ElseIf Mid(Text, ICNT, 1) = "u" Then
        Capitalize = Capitalize & "U"
    ElseIf Mid(Text, ICNT, 1) = "v" Then
        Capitalize = Capitalize & "V"
    ElseIf Mid(Text, ICNT, 1) = "z" Then
        Capitalize = Capitalize & "Z"
    ElseIf Mid(Text, ICNT, 1) = "ž" Then
        Capitalize = Capitalize & "Ž"
    ElseIf Mid(Text, ICNT, 1) = "x" Then
        Capitalize = Capitalize & "X"
    ElseIf Mid(Text, ICNT, 1) = "y" Then
        Capitalize = Capitalize & "Y"
    ElseIf Mid(Text, ICNT, 1) = "w" Then
        Capitalize = Capitalize & "W"
    ElseIf Mid(Text, ICNT, 1) = "ð" Then
        Capitalize = Capitalize & "Ð"
    ElseIf Mid(Text, ICNT, 1) = "ö" Then
        Capitalize = Capitalize & "Ö"
    ElseIf Mid(Text, ICNT, 1) = "ä" Then
        Capitalize = Capitalize & "Ä"
    ElseIf Mid(Text, ICNT, 1) = "ë" Then
        Capitalize = Capitalize & "Ë"
    ElseIf Mid(Text, ICNT, 1) = "ß" Then
        Capitalize = Capitalize & "ß"
    Else
        Capitalize = Capitalize & Mid(Text, ICNT, 1)
    End If
Next ICNT

End Function

Public Sub SelectIndex(index As Integer)
picSel.Visible = False
picBuff.Height = picBuffText.Height + 2
If index < 0 Then index = 0
If index > List1.ListCount - 1 Then
    index = List1.ListCount - 1
End If

If iSelected = index Then
Exit Sub
Else
iSelected = index
End If

picSel.Cls
picSel.Width = picOzadje.Width
picSel.Height = picBuff.Height

BitBlt picSel.hDC, 0, 0, picOzadje.Width, picBuff.Height, picOzadje.hDC, 0, picBuff.Height * index, SRCCOPY

picBuffText.Width = picBuffText.TextWidth(List1.List(index)) + 1
picBuffText.Cls
picBuffText.BackColor = SelBackColor
picBuffText.ForeColor = vbBlack
On Error Resume Next
'draw a rectagle around the caption
picSel.ForeColor = SelBorderColor
picSel.Line (Spacing * List2.List(index) + 19, 0)-(Spacing * List2.List(index) + 19, picSel.Height)
picSel.Line (Spacing * List2.List(index) + 20 + picBuffText.Width, 0)-(Spacing * List2.List(index) + 20 + picBuffText.Width, picSel.Height)
picSel.Line (Spacing * List2.List(index) + 19, 0)-(Spacing * List2.List(index) + 20 + picBuffText.Width, 0)
picSel.Line (Spacing * List2.List(index) + 19, picSel.Height - 1)-(Spacing * List2.List(index) + 20 + picBuffText.Width, picSel.Height - 1)

picBuffText.Print List1.List(index)

BitBlt picSel.hDC, Spacing * List2.List(index) + 20, 1, picBuffText.Width, picBuff.Height, picBuffText.hDC, 0, 0, SRCCOPY
picSel.Refresh

picSel.Top = index * picBuff.Height
picSel.Refresh

picOzadje.Refresh

picBuff.Cls
picBuff.Width = picBtn.Left - 2

BitBlt picBuff.hDC, 0, (picBuff.Height - IconBuff.Height) / 2, picBuff.Width, picBuff.Height, picOzadje.hDC, (List2.List(iSelected) * Spacing - 1), iSelected * picBuff.Height, SRCCOPY

picSel.Visible = True
'now copy the whole thing to the displayed picturebox
BitBlt PicDisplay.hDC, 2, (PicDisplay.Height - picBuff.Height) / 2, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
PicDisplay.Refresh

End Sub

Private Function CheckPath() As Boolean
On Error GoTo handle
CheckPath = True
Dir1.Path = Path ' JUST AN EASY WAY TO SEE IF THE PATH EXISTS
                 ' IF IT DOES NOT, DIR1 WILL RETURN AN ERROR

Exit Function
handle:
CheckPath = False
RaiseEvent ErrorOcured(vbError) 'vberror = 10 if the path is not found etc.

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    picBtn_Click
    
End If


End Sub

Private Sub UserControl_Resize()
If Not PicDisplay.Width = UserControl.Width / Screen.TwipsPerPixelX Then PicDisplay.Width = UserControl.Width / Screen.TwipsPerPixelX
If Not UserControl.Height = PicDisplay.Height * Screen.TwipsPerPixelY Then UserControl.Height = PicDisplay.Height * Screen.TwipsPerPixelY


End Sub


