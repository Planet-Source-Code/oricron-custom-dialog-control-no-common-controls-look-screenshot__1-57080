VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmDialog"
   ClientHeight    =   5025
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8745
   Icon            =   "Dialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OricronDialog.oList lstFiles 
      Left            =   6840
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.ListBox lstBack 
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstForward 
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picShdwV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   43
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picShdwH 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   42
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
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
      Left            =   0
      Picture         =   "Dialog.frx":000C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picExtensionsMnu 
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
      Height          =   495
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   38
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
      Begin VB.Label lblExtensionsMnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape shpExtensionsMnu 
         BorderColor     =   &H00000000&
         FillColor       =   &H8000000E&
         Height          =   1455
         Left            =   0
         Tag             =   "1"
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpExtensionsSel 
         BorderColor     =   &H00DBA573&
         FillColor       =   &H00F5E6D8&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   -1320
         Tag             =   "1"
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.PictureBox PicMnuViews 
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
      Height          =   1695
      Left            =   3840
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
      Begin OricronDialog.Button btnViewsMenu 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.Line lnMnuView 
         BorderColor     =   &H00FFFFFF&
         X1              =   56
         X2              =   88
         Y1              =   56
         Y2              =   56
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H80000010&
         FillColor       =   &H8000000E&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   120
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Timer tmrLeftMenu 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6120
      Top             =   3600
   End
   Begin VB.PictureBox PicMnuLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00996422&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   24
      Top             =   600
      Width           =   1335
      Begin VB.PictureBox picMnuLeftBackground 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   27
         Top             =   480
         Width           =   855
         Begin OricronDialog.Button btn 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   450
         End
      End
      Begin OricronDialog.Button btnUP 
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1080
         _ExtentX        =   1058
         _ExtentY        =   450
      End
      Begin OricronDialog.Button btnDOWN 
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   960
         Width           =   1200
         _ExtentX        =   1058
         _ExtentY        =   450
      End
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
      Left            =   7080
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBuff3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   583
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3810
      Width           =   8745
      Begin VB.PictureBox picExtensions 
         BackColor       =   &H00996422&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2280
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   525
         Width           =   3555
         Begin OricronDialog.Button btnExtensions 
            Height          =   135
            Left            =   3360
            TabIndex        =   37
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
         End
         Begin VB.PictureBox picEtensionsDisplay 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   101
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   120
            Width           =   1515
            Begin VB.Label lblExtensionsCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox picResize 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   8160
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   23
         Top             =   840
         Width           =   225
      End
      Begin VB.PictureBox picFile 
         BackColor       =   &H00996422&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2280
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Width           =   3555
         Begin VB.TextBox txtFile 
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
            Height          =   240
            Left            =   360
            TabIndex        =   22
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox picFileBack 
            BackColor       =   &H00996422&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   600
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   237
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   120
            Width           =   3555
         End
      End
      Begin VB.PictureBox picButtons 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   6720
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   78
         TabIndex        =   20
         Top             =   120
         Width           =   1170
         Begin VB.CommandButton cmdCancel 
            Caption         =   "cmdCancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "cmdOK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lblExtensions 
         AutoSize        =   -1  'True
         Caption         =   "lblExtensions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   32
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "lblFile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   30
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   8625
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   420
      Width           =   120
   End
   Begin VB.PictureBox PicMenuLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   0
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   420
      Width           =   1680
   End
   Begin VB.PictureBox PicMenuTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   583
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   8745
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   44
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox picToolbar 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4320
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   78
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1170
         Begin OricronDialog.Button mnuBtn 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   1058
            _ExtentY        =   450
         End
      End
      Begin OricronDialog.DirView DirView 
         Height          =   330
         Left            =   1800
         TabIndex        =   0
         Top             =   45
         Width           =   1935
         _extentx        =   3413
         _extenty        =   582
      End
      Begin VB.Label lblBrowse 
         AutoSize        =   -1  'True
         Caption         =   "lblBrowse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
   End
   Begin OricronDialog.FileView FileView 
      Height          =   1935
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   2535
      _extentx        =   4471
      _extenty        =   3413
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   8
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   7
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Image mnuIcon 
      Height          =   240
      Index           =   5
      Left            =   6600
      Picture         =   "Dialog.frx":04F6
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   4
      Left            =   5640
      Picture         =   "Dialog.frx":0A80
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   3
      Left            =   5640
      Picture         =   "Dialog.frx":0E0A
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   2
      Left            =   5640
      Picture         =   "Dialog.frx":1394
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   1
      Left            =   5640
      Picture         =   "Dialog.frx":191E
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   0
      Left            =   5640
      Picture         =   "Dialog.frx":1EA8
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DOWNImage 
      Height          =   180
      Left            =   6120
      Picture         =   "Dialog.frx":2432
      Top             =   2160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image UPImage 
      Height          =   180
      Left            =   6120
      Picture         =   "Dialog.frx":24E4
      Top             =   1800
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image mnuIcon 
      Height          =   360
      Index           =   4
      Left            =   6600
      Top             =   2040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image mnuIcon 
      Height          =   240
      Index           =   3
      Left            =   6600
      Picture         =   "Dialog.frx":2596
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image mnuIcon 
      Height          =   240
      Index           =   2
      Left            =   6600
      Picture         =   "Dialog.frx":2B20
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image mnuIcon 
      Height          =   240
      Index           =   1
      Left            =   6600
      Picture         =   "Dialog.frx":30AA
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image mnuIcon 
      Height          =   240
      Index           =   0
      Left            =   6600
      Picture         =   "Dialog.frx":3634
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim icq As Boolean
Dim tLeft As Integer
Dim K As Boolean
Dim Ban() As Integer
Dim DeskHdc&, ret&
Dim NoUpdate As Boolean
'AlphaBlending - to get blendet colors (buttons etc.)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)
Dim ColorBack As OLE_COLOR

Private Type ColorAndAlpha
    r                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
End Type

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Public iView0 As String
Public iView1 As String
Public iView2 As String
Public iView3 As String
Public iView4 As String
Public SaveMode As Boolean

Public iBtnBackIndex As Integer
Dim iBtnForwardIndex As Integer
Dim iBtnParentFolderIndex As Integer
Dim iBtnNewFolderIndex As Integer
Dim iBtnViewsIndex As Integer
Dim iBtnDeleteIndex As Integer

Dim DoNotClearUndo As Boolean

Dim MouseX As Integer
Dim MouseY As Integer

Dim Resizing As Boolean
Public iSizable As Boolean


Public MinWidth As Integer
Public MinHeight As Integer

Public ArrayButtons As Integer
Dim ArrayButtonPath(12) As String

Private Type POINTAPI
  X           As stdole.OLE_XPOS_PIXELS
  Y           As stdole.OLE_YPOS_PIXELS
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'To get icons from files...
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long

Dim iIcon As Long

'This code is written by Tim Misset - to get icons from folders i took some code from him - thanks;)
'Const LARGE_ICON As Integer = 32                '  do not need that so i removed it
'Const SMALL_ICON As Integer = 16                '  do not need that so i removed it
Const MAX_PATH = 260
Const ILD_TRANSPARENT = &H1                      '  Display transparent
Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Const SHGFI_EXETYPE = &H2000                     '  return exe type
Const SHGFI_LARGEICON = &H0                     '  get large icon - do not need that so i removed it
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

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Sub DrawPicMnuSep()
Dim ICNT As Integer
For ICNT = 1 To UBound(Ban)
    picToolbar.ForeColor = AlphaBlend(vb3DDKShadow, frmDialog.BackColor, 90)
    picToolbar.Line (Ban(ICNT) + 3, 5)-(Ban(ICNT) + 3, 19)
    
    picToolbar.ForeColor = AlphaBlend(vbWhite, frmDialog.BackColor, 190)
    picToolbar.Line (Ban(ICNT) + 4, 6)-(Ban(ICNT) + 4, 20)

Next ICNT


End Sub

Public Sub SetMnuExtensions()
On Error Resume Next
btnExtensions.Top = 1

picBuff.Height = picExtensions.Height - 2
picBuff.Width = 13

picBuff.BackColor = &H8000000F

picBuff.ForeColor = &H80000014
picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
picBuff.Line (0, 0)-(0, picBuff.Height - 1)

picBuff.ForeColor = &H80000010
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

DropDownIcon.BackColor = picBuff.BackColor

BitBlt picBuff.hDC, (picBuff.Width - DropDownIcon.Width) / 2, (picBuff.Height - DropDownIcon.Height) / 2, DropDownIcon.Width, DropDownIcon.hwnd, DropDownIcon.hDC, 0, 0, vbSrcCopy

Set btnExtensions.NormalImage = picBuff.Image
Set btnExtensions.FocusedImage = picBuff.Image

picBuff.Cls
picBuff.BackColor = &H8000000F

picBuff.ForeColor = &H80000010
picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
picBuff.Line (0, 0)-(0, picBuff.Height - 1)
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

DropDownIcon.BackColor = picBuff.BackColor

BitBlt picBuff.hDC, (picBuff.Width - DropDownIcon.Width) / 2 + 1, (picBuff.Height - DropDownIcon.Height) / 2 + 1, DropDownIcon.Width, DropDownIcon.hwnd, DropDownIcon.hDC, 0, 0, SRCCOPY

Set btnExtensions.PressedImage = picBuff.Image

picBuff.BackColor = vbWhite

Dim ICNT As Integer
Dim asd() As String
asd() = Split(dDialogFilter, "|")


For ICNT = 1 To lblExtensionsMnu.Count - 1
    Unload lblExtensionsMnu(ICNT)
Next ICNT

For ICNT = 1 To (UBound(asd) + 1) / 2
    Load lblExtensionsMnu(ICNT)
    lblExtensionsMnu(ICNT).Left = 2 * Screen.TwipsPerPixelX
    lblExtensionsMnu(ICNT).Top = (ICNT - 1) * lblExtensionsMnu(ICNT).Height + 2 * Screen.TwipsPerPixelY
    
    lblExtensionsMnu(ICNT).caption = asd((ICNT - 1) * 2)
    lblExtensionsMnu(ICNT).Tag = asd((ICNT - 1) * 2 + 1)
    lblExtensionsMnu(ICNT).Visible = True
Next ICNT

shpExtensionsSel.Left = 1 * Screen.TwipsPerPixelX

shpExtensionsSel.Height = lblExtensionsMnu(lblExtensionsMnu.Count - 1).Height + 2 * Screen.TwipsPerPixelY

shpExtensionsSel.ZOrder 1
SetWindowLong picExtensionsMnu.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
SetParent picExtensionsMnu.hwnd, 0
picExtensionsMnu.Height = ((lblExtensionsMnu.Count - 1) * lblExtensionsMnu(lblExtensionsMnu.Count - 1).Height) + 4 * Screen.TwipsPerPixelY
End Sub
Public Sub RedrawMnuExtensions()
btnExtensions.Top = 1

picBuff.Height = picExtensions.Height - 2
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

Set btnExtensions.NormalImage = picBuff.Image
Set btnExtensions.FocusedImage = picBuff.Image

picBuff.Cls
picBuff.BackColor = &H8000000F

picBuff.ForeColor = &H80000010
picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
picBuff.Line (0, 0)-(0, picBuff.Height - 1)
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

DropDownIcon.BackColor = picBuff.BackColor

BitBlt picBuff.hDC, (picBuff.Width - DropDownIcon.Width) / 2 + 1, (picBuff.Height - DropDownIcon.Height) / 2 + 1, DropDownIcon.Width, DropDownIcon.hwnd, DropDownIcon.hDC, 0, 0, SRCCOPY

Set btnExtensions.PressedImage = picBuff.Image

picBuff.BackColor = vbWhite

End Sub

Private Sub btn_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)

If btn(index).Top + btn(index).Height > btnDOWN.Top - picMnuLeftBackground.Top And btnDOWN.Visible = True Then ' we must set the height to bi just the visible part, because o setcapture!!!
    btn(index).Height = btnDOWN.Top - picMnuLeftBackground.Top - btn(index).Top
End If

If btn(index).Top + picMnuLeftBackground.Top < btnUP.Height Then
    btn(index).TopRegion = btnUP.Height - picMnuLeftBackground.Top - btn(index).Top
End If

End Sub

Private Sub btn_MouseOut(index As Integer)
btn(index).ResetSize
btn(index).TopRegion = 0

End Sub

Private Sub btn_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Button = vbLeftButton Then
    FileView.Path = Tag
    FileView.SetFocus
End If

End Sub

Private Sub btnDOWN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrLeftMenu.Tag = "DOWN"
tmrLeftMenu_Timer
tmrLeftMenu.Interval = 300
tmrLeftMenu.Enabled = True

End Sub

Private Sub btnDOWN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrLeftMenu.Enabled = False

End Sub

Private Sub btnExtensions_Click()
Dim Rec As RECT

GetWindowRect picExtensions.hwnd, Rec
Dim ICNT As Integer
Dim K As Integer
K = 0
For ICNT = 1 To lblExtensionsMnu.Count - 1
    If lblExtensionsMnu(ICNT).Width > K Then K = lblExtensionsMnu(ICNT).Width
Next ICNT

If K > (picExtensions.Width - 4) * Screen.TwipsPerPixelX Then
    picExtensionsMnu.Width = K + 4 * Screen.TwipsPerPixelX
Else
    picExtensionsMnu.Width = picExtensions.Width * Screen.TwipsPerPixelX
End If

If (Rec.Bottom) * Screen.TwipsPerPixelY + picExtensionsMnu.Height > Screen.Height - 40 * Screen.TwipsPerPixelY Then
    picExtensionsMnu.Top = (Rec.Top) * Screen.TwipsPerPixelY - picExtensionsMnu.Height
Else
    picExtensionsMnu.Top = (Rec.Bottom) * Screen.TwipsPerPixelY
End If
picExtensionsMnu.Left = Rec.Left * Screen.TwipsPerPixelX

picExtensionsMnu.ZOrder

picExtensionsMnu.Visible = True
picExtensionsMnu.ZOrder
SetCapture picExtensionsMnu.hwnd

End Sub

Private Sub btnUP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrLeftMenu.Tag = "UP"
tmrLeftMenu_Timer
tmrLeftMenu.Interval = 300
tmrLeftMenu.Enabled = True

End Sub

Private Sub btnUP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrLeftMenu.Enabled = False

End Sub

Private Sub cmdCancel_Click()
dCancel = True
lstFiles.Clear
Me.Hide

End Sub

Private Sub cmdOK_Click()

Dim ICNT As Integer
Dim cc As ListBox
Dim K As String
Dim sk As String

Set cc = FileView.GetSelectedFiles
Me.lstFiles.Clear

If SaveMode = False Then
    For ICNT = 0 To cc.ListCount - 1
        Me.lstFiles.AddItem cc.List(ICNT)
    Next ICNT
Else
    K = Me.FileView.Path
    If Right(K, 1) <> "\" Then K = K & "\"
    
    sk = Mid$(lblExtensionsMnu(shpExtensionsSel.Tag).Tag, 2)
    If Capitalize(Right(Me.txtFile.Text, Len(sk))) <> Capitalize(sk) Then
        Me.lstFiles.AddItem K & Me.txtFile.Text & sk
    Else
        Me.lstFiles.AddItem K & Me.txtFile.Text
    End If
    
End If

If lstFiles.ListCount > 0 Then
    dCancel = False
    Me.Hide
End If

End Sub

Private Sub DirView_Click(index As Integer, Path As String)
If FileView.Path <> Path And NoUpdate = False Then
    Me.FileView.Path = Path
    DirView.Path = Path
    FileView.SetFocus
End If

End Sub

Private Sub FileView_DirSelect(Path As String)
NoUpdate = True
DirView.Path = Path
DirView.Refresh
NoUpdate = False

End Sub

Private Sub FileView_FilePreSelect(Files As ListBox)
Dim ICNT As Integer
Dim s As String
If Files.ListCount = 0 Then
    frmDialog.txtFile.Text = ""
ElseIf Files.ListCount = 1 Then
    frmDialog.txtFile.Text = Files.List(0)
Else
    s = """" & Files.List(0) & """"
    
    For ICNT = 1 To Files.ListCount - 1
        s = s & " " & """" & Files.List(ICNT) & """"
    Next ICNT
    
    frmDialog.txtFile.Text = s
End If

End Sub


Private Sub FileView_FileSelect(Files As ListBox)
Dim ICNT As Integer
lstFiles.Clear
dCancel = False
For ICNT = 0 To Files.ListCount - 1
    lstFiles.AddItem Files.List(ICNT)
Next ICNT

Me.Hide
End Sub

Private Sub FileView_DeletableItemSelected(AnyDeletableItemSelected As Boolean)
If mnuBtn(iBtnDeleteIndex).cEnabled <> AnyDeletableItemSelected Then mnuBtn(iBtnDeleteIndex).Enabled = AnyDeletableItemSelected

End Sub

Private Sub FileView_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then MnuBtn_MouseUp iBtnParentFolderIndex, vbLeftButton, Shift, 1, 1, "#PARENTFOLDER"

End Sub

Private Sub FileView_MenuNewFolderClick()
MnuBtn_MouseUp 0, vbLeftButton, 0, 0, 0, "#NEWFOLDER"
End Sub

Private Sub FileView_PathChange(NewPath As String, OldPath As String)
lstBack.AddItem OldPath, 0
If mnuBtn(iBtnBackIndex).cEnabled = False Then mnuBtn(iBtnBackIndex).Enabled = True
If NewPath = DirView.DesktopFolder Then
    mnuBtn(iBtnParentFolderIndex).Enabled = False
Else
    If mnuBtn(iBtnParentFolderIndex).cEnabled = False Then mnuBtn(iBtnParentFolderIndex).Enabled = True
End If
If NewPath = "#MYCOMPUTER" Then
    mnuBtn(iBtnNewFolderIndex).Enabled = False
Else
    If mnuBtn(iBtnNewFolderIndex).cEnabled = False Then mnuBtn(iBtnNewFolderIndex).Enabled = True
End If

If DoNotClearUndo = False Then
    lstForward.Clear
    If mnuBtn(iBtnForwardIndex).cEnabled = True Then mnuBtn(iBtnForwardIndex).Enabled = False
Else

End If

DialogPath = NewPath
End Sub

Private Sub FileView_ViewChange(NewView As Integer)
If icq = False Then
    mnuIcon(4).Picture = MnuViewIcon(NewView).Picture
    setToolBar
    icq = True
End If

End Sub


Private Sub Form_Load()
Dim c         As ColorAndAlpha
OleTranslateColor Me.BackColor, 0, VarPtr(c)
ColorBack = RGB(c.r, c.G, c.B) 'we set this color to identify if the user changed colors

Me.Move Me.Left, Me.Top, MinWidth, MinHeight

picToolbar.Height = PicMenuTop.Height
picFile.Height = 22
picExtensions.Height = 22
picToolbar.Top = (PicMenuTop.Height - picToolbar.Height) / 2
PicMnuLeft.Top = PicMenuTop.Height

picResize.AutoRedraw = True

picResize.BackColor = picBottom.BackColor
picResize.ForeColor = AlphaBlend(AlphaBlend(vbButtonFace, vbWhite, 170), vb3DDKShadow, 70)
picFile.BackColor = AlphaBlend(AlphaBlend(vbButtonFace, vbWhite, 170), vb3DDKShadow, 70)
picExtensions.BackColor = AlphaBlend(AlphaBlend(vbButtonFace, vbWhite, 170), vb3DDKShadow, 70)

picResize.Line (11, 13)-(14, 10)
picResize.Line (12, 13)-(14, 11)

picResize.Line (9, 13)-(14, 8)
picResize.Line (8, 13)-(14, 7)

picResize.Line (6, 13)-(14, 5)
picResize.Line (5, 13)-(14, 4)

picResize.Refresh

setToolBar

SetLeftToolbar

End Sub

Public Sub setToolBar()

'On Error Resume Next
Dim ICNT As Integer
For ICNT = 1 To mnuBtn.Count - 1
    Unload mnuBtn(ICNT)
Next ICNT

tLeft = 7
picToolbar.BackColor = AlphaBlend(frmDialog.BackColor, vb3DHighlight, 170)
picToolbar.Height = 25

ReDim Ban(0)

DrawToolBar 1, "#BACK" '1 FOR BUTTON, 0 FOR SEPARATOR, AND A TAG FOR A SPECIFIC BUTTON... aDD MORE, BUT MODIFY THE DRAWTOOLBAR A BIT...
DrawToolBar 1, "#FORWARD"
DrawToolBar 0
DrawToolBar 1, "#PARENTFOLDER"
DrawToolBar 0
DrawToolBar 1, "#DELETE"
DrawToolBar 1, "#NEWFOLDER"
DrawToolBar 0
mnuIcon(4).Picture = MnuViewIcon(FileView.View)
DrawToolBar 1, "#MVIEW"


Me.picToolbar.Width = tLeft + 13
'to draw gradient behind the toolbar
Dim firsColor As OLE_COLOR
Dim SecondColor As OLE_COLOR

firsColor = vb3DHighlight
SecondColor = AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240)

picToolbar.AutoRedraw = True

picToolbar.Top = (PicMenuTop.Height - picToolbar.Height) / 2

For ICNT = 1 To Me.picToolbar.Height
    picToolbar.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * ICNT / picToolbar.Height))
    picToolbar.Line (0, ICNT - 1)-(picToolbar.Width, ICNT - 1)
    picToolbar.Refresh
Next ICNT

mnuBtn(iBtnBackIndex).Enabled = False
mnuBtn(iBtnForwardIndex).Enabled = False

picToolbar.ForeColor = AlphaBlend(vb3DHighlight, vb3DDKShadow, 100)
picToolbar.Line (picToolbar.Width - 11, 0)-(picToolbar.Width - 11, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 10, 0)-(picToolbar.Width - 10, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 9, 0)-(picToolbar.Width - 9, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 8, 0)-(picToolbar.Width - 8, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 7, 0)-(picToolbar.Width - 7, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 6, 0)-(picToolbar.Width - 6, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 5, 0)-(picToolbar.Width - 5, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 4, 0)-(picToolbar.Width - 4, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 3, 0)-(picToolbar.Width - 3, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 2, 1)-(picToolbar.Width - 2, picToolbar.Height - 1)
picToolbar.Line (picToolbar.Width - 1, 2)-(picToolbar.Width - 1, picToolbar.Height - 2)
picToolbar.Line (picToolbar.Width - 12, 0)-(picToolbar.Width - 12, 1)
picToolbar.Line (picToolbar.Width - 12, picToolbar.Height - 1)-(picToolbar.Width - 12, picToolbar.Height)

'draws soft edges
picToolbar.ForeColor = AlphaBlend(picToolbar.ForeColor, frmDialog.BackColor, 90)
picToolbar.Line (picToolbar.Width - 2, 0)-(picToolbar.Width, 2)
picToolbar.Line (picToolbar.Width - 2, picToolbar.Height - 1)-(picToolbar.Width, picToolbar.Height - 3)
picToolbar.Line (picToolbar.Width - 13, 0)-(picToolbar.Width - 11, 2)
picToolbar.Line (picToolbar.Width - 13, picToolbar.Height - 1)-(picToolbar.Width - 11, picToolbar.Height - 3)

picToolbar.ForeColor = AlphaBlend(vb3DHighlight, frmDialog.BackColor, 90)
picToolbar.Line (1, 0)-(0, 2)

picToolbar.ForeColor = frmDialog.BackColor
picToolbar.Line (picToolbar.Width - 1, 0)-(picToolbar.Width, 0)
picToolbar.Line (picToolbar.Width - 2, picToolbar.Height)-(picToolbar.Width, picToolbar.Height - 2)
picToolbar.Line (0, 0)-(0, 1)
picToolbar.Line (0, picToolbar.Height - 1)-(0, picToolbar.Height - 2)

picToolbar.ForeColor = AlphaBlend(vbWhite, frmDialog.BackColor, 190)

picToolbar.Line (4, 6)-(6, 6)
picToolbar.Line (4, 7)-(6, 7)

picToolbar.Line (4, 11)-(6, 11)
picToolbar.Line (4, 10)-(6, 10)

picToolbar.Line (4, 14)-(6, 14)
picToolbar.Line (4, 15)-(6, 15)

picToolbar.Line (4, 18)-(6, 18)
picToolbar.Line (4, 19)-(6, 19)


picToolbar.ForeColor = AlphaBlend(vb3DDKShadow, frmDialog.BackColor, 105)

picToolbar.Line (3, 6)-(5, 6)
picToolbar.Line (3, 5)-(5, 5)

picToolbar.Line (3, 9)-(5, 9)
picToolbar.Line (3, 10)-(5, 10)

picToolbar.Line (3, 14)-(5, 14)
picToolbar.Line (3, 13)-(5, 13)

picToolbar.Line (3, 18)-(5, 18)
picToolbar.Line (3, 17)-(5, 17)

'draws separators between buttons
DrawPicMnuSep

End Sub

Public Sub RedrawToolBar()
Dim ICNT As Integer

picToolbar.BackColor = AlphaBlend(frmDialog.BackColor, vb3DHighlight, 170)

'to draw gradient behind the toolbar
Dim firsColor As OLE_COLOR
Dim SecondColor As OLE_COLOR

firsColor = vb3DHighlight
SecondColor = AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240)

picToolbar.AutoRedraw = True

picToolbar.Top = (PicMenuTop.Height - picToolbar.Height) / 2

For ICNT = 1 To Me.picToolbar.Height
    picToolbar.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * ICNT / picToolbar.Height))
    picToolbar.Line (0, ICNT - 1)-(picToolbar.Width, ICNT - 1)
    picToolbar.Refresh
Next ICNT

For ICNT = 1 To mnuBtn.Count - 1
    DrawBtnMnu ICNT, mnuBtn(ICNT).sTag
Next ICNT

picToolbar.ForeColor = AlphaBlend(vb3DHighlight, vb3DDKShadow, 100)
picToolbar.Line (picToolbar.Width - 11, 0)-(picToolbar.Width - 11, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 10, 0)-(picToolbar.Width - 10, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 9, 0)-(picToolbar.Width - 9, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 8, 0)-(picToolbar.Width - 8, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 7, 0)-(picToolbar.Width - 7, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 6, 0)-(picToolbar.Width - 6, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 5, 0)-(picToolbar.Width - 5, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 4, 0)-(picToolbar.Width - 4, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 3, 0)-(picToolbar.Width - 3, picToolbar.Height)
picToolbar.Line (picToolbar.Width - 2, 1)-(picToolbar.Width - 2, picToolbar.Height - 1)
picToolbar.Line (picToolbar.Width - 1, 2)-(picToolbar.Width - 1, picToolbar.Height - 2)
picToolbar.Line (picToolbar.Width - 12, 0)-(picToolbar.Width - 12, 1)
picToolbar.Line (picToolbar.Width - 12, picToolbar.Height - 1)-(picToolbar.Width - 12, picToolbar.Height)

'draws soft edges
picToolbar.ForeColor = AlphaBlend(picToolbar.ForeColor, frmDialog.BackColor, 90)
picToolbar.Line (picToolbar.Width - 2, 0)-(picToolbar.Width, 2)
picToolbar.Line (picToolbar.Width - 2, picToolbar.Height - 1)-(picToolbar.Width, picToolbar.Height - 3)
picToolbar.Line (picToolbar.Width - 13, 0)-(picToolbar.Width - 11, 2)
picToolbar.Line (picToolbar.Width - 13, picToolbar.Height - 1)-(picToolbar.Width - 11, picToolbar.Height - 3)

picToolbar.ForeColor = AlphaBlend(vb3DHighlight, frmDialog.BackColor, 90)
picToolbar.Line (1, 0)-(0, 2)

picToolbar.ForeColor = frmDialog.BackColor
picToolbar.Line (picToolbar.Width - 1, 0)-(picToolbar.Width, 0)
picToolbar.Line (picToolbar.Width - 2, picToolbar.Height)-(picToolbar.Width, picToolbar.Height - 2)
picToolbar.Line (0, 0)-(0, 1)
picToolbar.Line (0, picToolbar.Height - 1)-(0, picToolbar.Height - 2)

picToolbar.ForeColor = AlphaBlend(vbWhite, frmDialog.BackColor, 190)

picToolbar.Line (4, 6)-(6, 6)
picToolbar.Line (4, 7)-(6, 7)

picToolbar.Line (4, 11)-(6, 11)
picToolbar.Line (4, 10)-(6, 10)

picToolbar.Line (4, 14)-(6, 14)
picToolbar.Line (4, 15)-(6, 15)

picToolbar.Line (4, 18)-(6, 18)
picToolbar.Line (4, 19)-(6, 19)


picToolbar.ForeColor = AlphaBlend(vb3DDKShadow, frmDialog.BackColor, 105)

picToolbar.Line (3, 6)-(5, 6)
picToolbar.Line (3, 5)-(5, 5)

picToolbar.Line (3, 9)-(5, 9)
picToolbar.Line (3, 10)-(5, 10)

picToolbar.Line (3, 14)-(5, 14)
picToolbar.Line (3, 13)-(5, 13)

picToolbar.Line (3, 18)-(5, 18)
picToolbar.Line (3, 17)-(5, 17)

'draws separators between buttons
DrawPicMnuSep

End Sub

Private Sub DrawToolBar(object As Integer, Optional Tag As String)
On Error Resume Next
If object = 0 Then
    SetPicMnuSep
Else
    Load mnuBtn(mnuBtn.Count)
    DrawBtnMnu mnuBtn.Count - 1, Tag
    mnuBtn(mnuBtn.Count - 1).Left = tLeft
    mnuBtn(mnuBtn.Count - 1).Top = 1
    tLeft = tLeft + mnuBtn(mnuBtn.Count - 1).Width
    mnuBtn(mnuBtn.Count - 1).Visible = True
End If

End Sub

Private Sub SetPicMnuSep()
ReDim Preserve Ban(UBound(Ban) + 1)
Ban(UBound(Ban)) = tLeft

tLeft = tLeft + 7

End Sub

Private Sub DrawBtnMnu(index As Integer, Tag As String)
Dim iIndex As Integer
Dim cc As cMemDC

Select Case Tag
    Case "#BACK"
        iIndex = 0
        iBtnBackIndex = index
    Case "#PARENTFOLDER"
        iIndex = 1
        iBtnParentFolderIndex = index
    Case "#NEWFOLDER"
        iIndex = 2
        iBtnNewFolderIndex = index
    Case "#FORWARD"
        iIndex = 3
        iBtnForwardIndex = index
    Case "#MVIEW"
        iIndex = 4
        iBtnViewsIndex = index
    Case "#DELETE"
        iIndex = 5
        iBtnDeleteIndex = index
End Select

mnuBtn(index).SetTag Tag

If mnuBtn(index).cEnabled = True Then
    picBuff.Cls
    picBuff.BackColor = picToolbar.BackColor
    
    picBuff.Height = 22
    
    If iIndex = 4 Then
        picBuff.Width = 32
    Else
        picBuff.Width = 22
    End If
    Dim firsColor As OLE_COLOR
    Dim SecondColor As OLE_COLOR
    
    firsColor = vb3DHighlight
    SecondColor = AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240)
    
    Dim ICNT As Integer
    For ICNT = 1 To Me.picToolbar.Height
        picBuff.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * ICNT / picToolbar.Height))
        picBuff.Line (0, ICNT - 2)-(picBuff.Width, ICNT - 2)
        picBuff.Refresh
    Next ICNT
    
    picBuff3.AutoRedraw = True
    picBuff3.Cls

    picBuff3.AutoSize = True
    picBuff3.BackColor = RGB(255, 0, 255)
    'BitBlt picBuff3.hDC, 0, 0, 22, 22, picToolbar.hDC, 0, (picToolbar.Height - 16) / 2, SRCCOPY
    picBuff3.Picture = mnuIcon(iIndex)
        
    Set cc = DimBitmap(picBuff3.Picture, picBuff3.BackColor)
    'picBuff.Cls
    cc.BitBlt picBuff.hDC, (picBuff.Height - picBuff3.Height) / 2, (picBuff.Height - picBuff3.Height) / 2, 16, 16, 0, 0
    
    'BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    'picBuff.Refresh
    If iIndex = 4 Then
        picBuff.ForeColor = vbBlack
        picBuff.Line (picBuff.Width - 8, 10)-(picBuff.Width - 3, 10)
        picBuff.Line (picBuff.Width - 7, 11)-(picBuff.Width - 4, 11)
        picBuff.Line (picBuff.Width - 6, 12)-(picBuff.Width - 5, 12)
    End If
    
    Set mnuBtn(index).NormalImage = picBuff.Image
        
    If Tag = "#MVIEW" And PicMnuViews.Visible = True Then
        picBuff.BackColor = vb3DHighlight 'picToolbar.BackColor
        picBuff.ForeColor = shpBorder.BorderColor
        
        picBuff.Line (0, 0)-(picBuff.Width, 0)
        picBuff.Line (0, 0)-(0, picBuff.Height)
        picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)

        picBuff.ForeColor = vbBlack
        picBuff.Line (picBuff.Width - 8, 10)-(picBuff.Width - 3, 10)
        picBuff.Line (picBuff.Width - 7, 11)-(picBuff.Width - 4, 11)
        picBuff.Line (picBuff.Width - 6, 12)-(picBuff.Width - 5, 12)
    Else
        picBuff.ForeColor = vbHighlight
        picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)

        picBuff.Line (0, 0)-(picBuff.Width, 0)
        picBuff.Line (0, 0)-(0, picBuff.Height)
        picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
        picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
        
        If iIndex = 4 Then
            picBuff.Line (picBuff.Width - 11, 0)-(picBuff.Width - 11, picBuff.Height)
            
            picBuff.ForeColor = vbBlack
            picBuff.Line (picBuff.Width - 8, 10)-(picBuff.Width - 3, 10)
            picBuff.Line (picBuff.Width - 7, 11)-(picBuff.Width - 4, 11)
            picBuff.Line (picBuff.Width - 6, 12)-(picBuff.Width - 5, 12)
        End If
    End If
    
    
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.Picture = mnuIcon(iIndex)
    
    BitBlt picBuff.hDC, (picBuff.Height - picBuff3.Height) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    Set mnuBtn(index).FocusedImage = picBuff.Image
    If Tag = "#MVIEW" And PicMnuViews.Visible = True Then
        Set mnuBtn(index).PressedImage = picBuff.Image
    End If
        
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    

    picBuff3.BackColor = picBuff.BackColor
    picBuff3.Picture = mnuIcon(iIndex)
    
    BitBlt picBuff.hDC, (picBuff.Height - picBuff3.Height) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    If iIndex = 4 Then
        picBuff.Line (picBuff.Width - 11, 0)-(picBuff.Width - 11, picBuff.Height)
        
        picBuff.ForeColor = vbBlack
        picBuff.Line (picBuff.Width - 8, 10)-(picBuff.Width - 3, 10)
        picBuff.Line (picBuff.Width - 7, 11)-(picBuff.Width - 4, 11)
        picBuff.Line (picBuff.Width - 6, 12)-(picBuff.Width - 5, 12)
    End If
    
    
    If Tag = "#MVIEW" And PicMnuViews.Visible = True Then
    Else
        Set mnuBtn(index).PressedImage = picBuff.Image
    End If
picBuff.Picture = LoadPicture("")
'm_clrDisabledMenuBorder = vbButtonShadow
'm_clrDisabledMenuBack = pvAlphaBlend(m_clrMenuBack, vbWindowBackground, 128)
'm_clrDisabledMenuFore = vbGrayText
Else
    picBuff.Cls
    picBuff.BackColor = picToolbar.BackColor
    picBuff.Width = 22
    picBuff.Height = 22
    
    
    firsColor = vb3DHighlight
    SecondColor = AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240)

    For ICNT = 1 To Me.picToolbar.Height
        picBuff.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * ICNT / picToolbar.Height))
        picBuff.Line (0, ICNT - 2)-(picBuff.Width, ICNT - 2)
        picBuff.Refresh
    Next ICNT
    picBuff3.AutoSize = True
    
    picBuff3.BackColor = picToolbar.BackColor
    
    picBuff3.Picture = mnuIcon(iIndex)
    K = True
    Set cc = DisabledPicture(picBuff3.Picture, picBuff3.BackColor, pvGetLuminance(picToolbar.BackColor), True)
    
    cc.BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, 16, 16, 0, 0
    
    'BitBlt picBuff.hDC, (picBuff.Width - PicBuff2.Width) / 2, (picBuff.Height - PicBuff2.Height) / 2, PicBuff2.Width, PicBuff2.Height, PicBuff2.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    K = False
    Set mnuBtn(index).NormalImage = picBuff.Image
    
    picBuff.ForeColor = AlphaBlend(vb3DDKShadow, vb3DHighlight, 90)
    picBuff.BackColor = AlphaBlend(picBuff.ForeColor, vb3DHighlight, 70)
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    
    picBuff3.BackColor = picBuff.BackColor
    
    picBuff3.Picture = mnuIcon(iIndex)
    
    Set cc = DisabledPicture(picBuff3.Picture, picBuff3.BackColor, pvGetLuminance(picToolbar.BackColor))
    cc.BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, 16, 16, 0, 0
    
    'BitBlt picBuff.hDC, (picBuff.Width - PicBuff2.Width) / 2, (picBuff.Height - PicBuff2.Height) / 2, PicBuff2.Width, PicBuff2.Height, PicBuff2.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set mnuBtn(index).FocusedImage = picBuff.Image
    
    picBuff.ForeColor = AlphaBlend(vb3DDKShadow, vb3DHighlight, 90)
    picBuff.BackColor = AlphaBlend(picBuff.ForeColor, vb3DHighlight, 70)
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    
    picBuff3.BackColor = picBuff.BackColor
    
    picBuff3.Picture = mnuIcon(iIndex)
    
    Set cc = DisabledPicture(picBuff3.Picture, picBuff3.BackColor, pvGetLuminance(picToolbar.BackColor))
    cc.BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, 16, 16, 0, 0
    
    picBuff.Refresh
    
    Set mnuBtn(index).PressedImage = picBuff.Image

End If

End Sub

Private Function DisabledPicture(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR, ByVal TresholdLuminance As Long, Optional lighten As Boolean) As cMemDC
    Dim lI              As Long
    Dim lJ              As Long
    Dim lK              As Long
    Dim iForeColor      As ColorAndAlpha
    
    Set DisabledPicture = New cMemDC
    Dim c As Long
    With DisabledPicture
        .Init 16, 16
        .Cls MaskColor
        .PaintPicture oPic, 0, 0, 17, 17, -10, -10, vbSrcCopy, clrMask:=MaskColor
        For lJ = 0 To 16
            For lI = 0 To 16
                lK = .GetPixel(lI, lJ)
                If lK <> MaskColor Then
                    OleTranslateColor lK, 0, VarPtr(iForeColor)
                    'gets the gray value for the picture and then lighten it a bit (+15)
                    c = iForeColor.r
                    c = c + iForeColor.G
                    c = c + iForeColor.B
                    c = c / 3
                    
                    If lighten = True Then c = c + 35
                    
                    If c > 255 Then c = 255
                    
                    .SetPixel lI, lJ, AlphaBlend(MaskColor, RGB(c, c, c), 70)
                Else
                    If K = True Then .SetPixel lI, lJ, AlphaBlend(AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240), vb3DHighlight, Int(255 * (5 + lJ) / picToolbar.Height))
                End If
            Next
        Next
    End With
End Function

Private Function pvGetLuminance(ByVal clrColor As Long) As Long
    Dim rgbColor        As ColorAndAlpha
    
    OleTranslateColor clrColor, 0, VarPtr(rgbColor)
    pvGetLuminance = (rgbColor.r * 76& + rgbColor.G * 150& + rgbColor.B * 29&) \ 255&

End Function

Private Sub Form_Resize()
picToolbar.Left = Me.ScaleWidth - picToolbar.Width - picRight.Width - 1

picButtons.Left = Me.ScaleWidth - picButtons.Width - picRight.Width

PicMnuLeft.Height = Me.ScaleHeight - PicMenuTop.Height - 15

'Resize picture position (the corner)
picResize.Left = Me.ScaleWidth - picResize.Width
picResize.Top = picBottom.Height - picResize.Height

'Browse label position
If frmDialog.lblBrowse.Width > PicMenuLeft.Width - 5 Then
    lblBrowse.Left = 5
    DirView.Left = 5 + lblBrowse.Width
Else
    lblBrowse.Left = PicMenuLeft.Width - lblBrowse.Width
    DirView.Left = PicMenuLeft.Width
End If

'File and extensions label positions
Dim xLeft As Integer
If lblFile.Width > lblExtensions.Width Then
    If lblFile.Width > PicMenuLeft.Width - 5 Then
        xLeft = PicMenuLeft.Width + lblFile.Width + 5
    Else
        xLeft = PicMenuLeft.Width * 2 - 5
    End If
Else
    If lblExtensions.Width > PicMenuLeft.Width - 5 Then
        xLeft = PicMenuLeft.Width + lblExtensions.Width + 5
    Else
        xLeft = PicMenuLeft.Width * 2 - 5
    End If
End If

picFile.Left = xLeft
picExtensions.Left = xLeft

lblFile.Left = xLeft - 5 - lblFile.Width
lblExtensions.Left = xLeft - 5 - lblExtensions.Width

lblFile.Top = picFile.Top + 4
lblExtensions.Top = picExtensions.Top + 4

On Error Resume Next
picFile.Width = picButtons.Left - picFile.Left - 5
picExtensions.Width = picButtons.Left - picExtensions.Left - 5

DirView.Width = picToolbar.Left - DirView.Left - 3

FileView.Top = PicMenuTop.Height
FileView.Left = PicMenuLeft.Width

If FileView.Height <> Me.ScaleHeight - FileView.Top - picBottom.Height Or FileView.Width <> Me.ScaleWidth - FileView.Left - picRight.Width Then
    FileView.Move FileView.Left, FileView.Top, Me.ScaleWidth - FileView.Left - picRight.Width, Me.ScaleHeight - FileView.Top - picBottom.Height
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
Me.Hide

DialogView = FileView.View

dCancel = True
DialogWidth = Me.Width
DialogHeight = Me.Height

If DoNotUnload = True Then
    Cancel = 1
Else
    Cancel = 0
End If

End Sub

Private Sub mnuBtn_Click(index As Integer)
If index = iBtnViewsIndex Then
    'If PicMnuViews.visible = False Then
        SetCapture PicMnuViews.hwnd
    'End If
End If

End Sub

Private Sub mnuBtn_EnableStateChange(index As Integer, Enabled As Boolean)
DrawBtnMnu index, mnuBtn(index).GetTag(0)

End Sub

Private Sub mnuBtn_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Tag = "#MVIEW" Then
If index = iBtnViewsIndex Then
    If PicMnuViews.Visible = False Then
        mnuBtn(iBtnViewsIndex).SetState "H"
        SetMenuViews
           
        PicMnuViews.Visible = True
        picShdwH.Visible = True
        picShdwV.Visible = True
        
        DrawBtnMnu iBtnViewsIndex, mnuBtn(iBtnViewsIndex).sTag
        
    End If
End If
End If
End Sub

Private Sub MnuBtn_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Button = vbLeftButton Then
    On Error Resume Next
    Select Case Tag
        Case "#PARENTFOLDER"
            If DirView.Path = DirView.DesktopFolder Then ' if we're on Desktop we cannot go back anymore
            
            Else
                If Left(DirView.Path, 1) <> "#" And DirView.Path <> DirView.MyDocumentsFolder Then ' the root folder of MyComputer and MyDocuments will be Desktop, for others we move down the root
                    If Len(DirView.Path) > 3 Then
                        Dim ipath As String
                        
                        ipath = Left(DirView.Path, InStrRev(DirView.Path, "\") - 1)
                        If Len(ipath) < 3 Then ipath = ipath & "\"
                        FileView.Path = ipath
                        
                    Else
                        FileView.Path = "#MYCOMPUTER"
                        
                    End If
                Else
                    FileView.Path = DirView.DesktopFolder
            End If
            End If
        Case "#BACK"
            DoNotClearUndo = True
            lstForward.AddItem FileView.Path, 0
            If mnuBtn(iBtnForwardIndex).cEnabled = False Then mnuBtn(iBtnForwardIndex).Enabled = True
            FileView.Path = lstBack.List(0)
            lstBack.RemoveItem 0
            lstBack.RemoveItem 0
            
            If lstBack.ListCount = 0 Then
                mnuBtn(iBtnBackIndex).Enabled = False
            End If
            
            DoNotClearUndo = False
        Case "#FORWARD"
            DoNotClearUndo = True
            lstBack.AddItem FileView.Path, 0
            If mnuBtn(iBtnBackIndex).cEnabled = False Then mnuBtn(iBtnBackIndex).Enabled = True
            FileView.Path = lstForward.List(0)
            lstBack.RemoveItem 0
            lstForward.RemoveItem 0
            'lstBack.RemoveItem 0
            
            If lstForward.ListCount = 0 Then
                mnuBtn(iBtnForwardIndex).Enabled = False
            End If
            
            DoNotClearUndo = False
            
        Case "#NEWFOLDER"
            frmMKDir.Show vbModal
        Case "#DELETE"
            FileView.DeleteSelected
            
    End Select
End If

End Sub

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

Private Function DimBitmap(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR) As cMemDC
    Dim lI              As Long
    Dim lJ              As Long
    
    Set DimBitmap = New cMemDC
    With DimBitmap
        .Init 16, 16
        .Cls MaskColor
        .PaintPicture oPic, 0, 0, 17, 17, -10, -10, vbSrcCopy, clrMask:=MaskColor
        For lJ = 0 To 16
            For lI = 0 To 16
                If .GetPixel(lI, lJ) <> MaskColor Then
                    .SetPixel lI, lJ, AlphaBlend(picToolbar.BackColor, .GetPixel(lI, lJ), 70)
                Else
                    .SetPixel lI, lJ, AlphaBlend(AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240), vb3DHighlight, Int(255 * (5 + lJ) / picToolbar.Height))
                End If
            Next
        Next
    End With
End Function

Private Function DrawShadow(ByVal oPic As StdPicture, Vertical As Boolean) As cMemDC
Dim lI              As Long
Dim lJ              As Long

Set DrawShadow = New cMemDC
With DrawShadow
    
    Dim cc As PictureBox
    If Vertical = False Then
        Set cc = picShdwH
    Else
        Set cc = picShdwV
    End If
    
    .Init cc.Width, cc.Height
    .Cls
    .PaintPicture oPic, 0, 0, cc.Width, cc.Height, 0, 0, vbSrcCopy, clrMask:=RGB(255, 0, 255)

    For lJ = 0 To cc.Height - 1
        For lI = 0 To cc.Width - 1
            If Vertical = False Then
                If lI = 0 And lJ = 0 Then 'ALL PARTS BESSIDE 'ELSE' DRAWS SOFT EDGES
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                ElseIf lI = 1 And lJ = 1 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                ElseIf lI = 2 And lJ = 2 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                ElseIf lI = 3 And lJ = 3 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                ElseIf lI = 1 And lJ = 0 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 2 / 4 * 60)
                ElseIf lI = 2 And lJ = 1 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 2 / 4 * 60)
                ElseIf lI = 3 And lJ = 2 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 2 / 4 * 60)
                ElseIf lI = 2 And lJ = 0 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 3 / 4 * 60)
                ElseIf lI = 3 And lJ = 1 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 3 / 4 * 60)
                ElseIf lI = 3 And lJ = 0 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 * 60)
                ElseIf lI = 0 And lJ = 1 Then
                    .SetPixel lI, lJ, .GetPixel(lI, lJ)
                ElseIf lI = 0 And lJ = 2 Then
                    .SetPixel lI, lJ, .GetPixel(lI, lJ)
                ElseIf lI = 0 And lJ = 3 Then
                    .SetPixel lI, lJ, .GetPixel(lI, lJ)
                ElseIf lI = 1 And lJ = 2 Then
                    .SetPixel lI, lJ, .GetPixel(lI, lJ)
                ElseIf lI = 1 And lJ = 3 Then
                    .SetPixel lI, lJ, .GetPixel(lI, lJ)
                ElseIf lI = 2 And lJ = 3 Then
                    .SetPixel lI, lJ, .GetPixel(lI, lJ)
                Else
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), (cc.Height - lJ) / cc.Height * 60)
                End If
            Else
                'Exit Function
                If lJ = 3 And lI = 3 Then 'ALL PARTS BESSIDE 'ELSE' DRAWS SOFT EDGES
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 1 / 4 * 60)
                ElseIf lJ = 2 And lI = 2 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 1 / 4 * 60)
                ElseIf lJ = 1 And lI = 1 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 1 / 4 * 60)
                ElseIf lJ = 0 And lI = 0 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 1 / 4 * 60)
                ElseIf lJ = 2 And lI = 3 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 2 / 4 * 60)
                ElseIf lJ = 1 And lI = 2 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 2 / 4 * 60)
                ElseIf lJ = 0 And lI = 1 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 2 / 4 * 60)
                ElseIf lJ = 1 And lI = 3 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 3 / 4 * 60)
                ElseIf lJ = 0 And lI = 2 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 3 / 4 * 60)
                ElseIf lJ = 0 And lI = 3 Then
                    .SetPixel lJ, lI, AlphaBlend(vbBlack, .GetPixel(lJ, lI), 60)
                
                ElseIf lJ = cc.Height - 3 And lI = cc.Width - 1 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                ElseIf lJ = cc.Height - 2 And lI = cc.Width - 2 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                ElseIf lJ = cc.Height - 1 And lI = cc.Width - 3 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 1 / 4 * 60)
                    
                ElseIf lJ = cc.Height - 3 And lI = cc.Width - 2 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 2 / 4 * 60)
                ElseIf lJ = cc.Height - 2 And lI = cc.Width - 3 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 2 / 4 * 60)
                ElseIf lJ = cc.Height - 3 And lI = cc.Width - 3 Then
                    .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), 3 / 4 * 60)
                    

                Else
                   If lJ > 3 And lJ < cc.Height - 3 Then .SetPixel lI, lJ, AlphaBlend(vbBlack, .GetPixel(lI, lJ), (cc.Width - lI) / cc.Width * 60)
                End If
            End If
        Next
    Next
End With
    
End Function
Private Function DimBitmap2(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR, IsSelected As Boolean) As cMemDC
    Dim lI              As Long
    Dim lJ              As Long
    
    Set DimBitmap2 = New cMemDC
    With DimBitmap2
        .Init 16, 16
        .Cls MaskColor
        .PaintPicture oPic, 0, 0, 17, 17, -10, -10, vbSrcCopy, clrMask:=MaskColor
        For lJ = 0 To 16
            For lI = 0 To 16
                If .GetPixel(lI, lJ) <> MaskColor Then
                    .SetPixel lI, lJ, AlphaBlend(picToolbar.BackColor, .GetPixel(lI, lJ), 70)
                Else
                    If IsSelected = True Then
                        .SetPixel lI, lJ, MaskColor
                    Else
                        .SetPixel lI, lJ, AlphaBlend(AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240), vb3DHighlight, Int(255 * (5 + lI) / picToolbar.Height))
                    End If
                End If
            Next
        Next
    End With
    
End Function
Private Function GetHoverImage(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR) As cMemDC
    Dim lI              As Long
    Dim lJ              As Long
    
    Set GetHoverImage = New cMemDC
    With GetHoverImage
        .Init 16 + 2, 16 + 2
        .Cls MaskColor
        .PaintPicture oPic, 0, 0, 16, 16, -10, -10, vbSrcCopy, clrMask:=MaskColor
        For lJ = 19 To 2 Step -1
            For lI = 19 To 2 Step -1
                If .GetPixel(lI - 2, lJ - 2) <> MaskColor And .GetPixel(lI, lJ) = MaskColor Then
                    .SetPixel lI, lJ, AlphaBlend(picBuff.BackColor, picBuff.ForeColor, 70)
                End If
            Next
        Next
        
        For lJ = 0 To 18
            For lI = 0 To 18
                If .GetPixel(lI, lJ) = MaskColor Then
                     .SetPixel lI, lJ, picBuff.BackColor
                End If
            Next
        Next
    End With
End Function

Private Sub mnuBtn_TotalMouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If PicMnuViews.Visible = True Then
    Dim ix As Integer
    Dim iy As Integer
    
    ix = X
    iy = Y
    
    If ix < 0 Or iy < 0 Or iy > mnuBtn(iBtnViewsIndex).Height Or ix > mnuBtn(iBtnViewsIndex).Width Then
        PicMnuViews_MouseDown 0, 0, -1, -1
    End If
End If

End Sub

Private Sub picBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picBottom.MousePointer <> 0 Then Resizing = True
MouseX = picBottom.Width - X + (Me.Width / Screen.TwipsPerPixelX - Me.ScaleWidth) / 2
MouseY = picBottom.Height - Y + (Me.Width / Screen.TwipsPerPixelX - Me.ScaleWidth) / 2

End Sub


Private Sub picBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.iSizable = False Then Exit Sub

Dim w As Integer
Dim h As Integer
Dim c As POINTAPI
GetCursorPos c

If Button = vbLeftButton Then
    If Resizing = True Then
        w = frmDialog.Width
        h = frmDialog.Height
        
        If picBottom.MousePointer = 9 Or picBottom.MousePointer = 8 Then
            w = c.X * Screen.TwipsPerPixelX - frmDialog.Left + MouseX * Screen.TwipsPerPixelX
            If w < MinWidth Then
                w = MinWidth
            End If
        End If
        
        If picBottom.MousePointer = 7 Or picBottom.MousePointer = 8 Then
            h = c.Y * Screen.TwipsPerPixelY - frmDialog.Top + MouseY * Screen.TwipsPerPixelY
            If h < MinHeight Then
                h = MinHeight
            End If
        End If

        If frmDialog.Width <> w Or frmDialog.Height <> h Then
            frmDialog.Move frmDialog.Left, frmDialog.Top, w, h
            frmDialog.Refresh
        End If
        
    End If
Else
    If X >= picBottom.Width - 3 And Y <= picBottom.Height - 15 Then
        picBottom.MousePointer = 9
    ElseIf Y >= picBottom.Height - 3 And X <= picBottom.Width - 15 Then
        picBottom.MousePointer = 7
    ElseIf Y >= picBottom.Height - 15 And X >= picBottom.Width - 15 Then
        picBottom.MousePointer = 8
    Else
        picBottom.MousePointer = 0
    End If
End If

End Sub


Private Sub picBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Resizing = False

End Sub

Private Sub picExtensions_Click()
btnExtensions_Click

End Sub

Private Sub picExtensions_Resize()
picEtensionsDisplay.Top = 1
picEtensionsDisplay.Height = picExtensions.Height - 2
picEtensionsDisplay.Left = 1
picEtensionsDisplay.Width = picExtensions.Width - 2 - btnExtensions.Width
btnExtensions.Left = picEtensionsDisplay.Width + 1
lblExtensionsCap.Left = 1
lblExtensionsCap.Top = (picEtensionsDisplay.Height - lblExtensionsCap.Height) / 2

End Sub


Private Sub picExtensionsMnu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Rec As RECT, ok, lx, ly

ok = GetClientRect(picExtensionsMnu.hwnd, Rec)

lx = CLng(X)
ly = CLng(Y)
On Error Resume Next
If PtInRect(Rec, lx, ly) = 0 Then  'Returns 1 if true, 0 if false
    If shpExtensionsSel.Tag <> "" And shpExtensionsSel.Visible = True Then
        FileView.filter = lblExtensionsMnu(shpExtensionsSel.Tag).Tag
        lstBack.RemoveItem 0
        If lstBack.ListCount = 0 Then mnuBtn(iBtnBackIndex).Enabled = False
        frmDialog.lblExtensionsCap.caption = lblExtensionsMnu(shpExtensionsSel.Tag).caption
    End If
    
    picExtensionsMnu.Visible = False
    ReleaseCapture
Else
    picExtensionsMnu.Visible = False
    ReleaseCapture
End If

End Sub


Private Sub picExtensionsMnu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X >= 0 And X <= picExtensionsMnu.Width And Y >= 0 And Y <= picExtensionsMnu.Height Then
    If Not shpExtensionsSel.Width = picExtensionsMnu.Width - 2 * Screen.TwipsPerPixelX Then shpExtensionsSel.Width = picExtensionsMnu.Width - 2 * Screen.TwipsPerPixelX
    
    Dim K As Integer
    
    K = Int((Y - 2 * Screen.TwipsPerPixelY) / lblExtensionsMnu(1).Height)
    
    If K >= 0 And K < lblExtensionsMnu.Count - 1 Then
        If Not shpExtensionsSel.Top = K * lblExtensionsMnu(1).Height + 1 * Screen.TwipsPerPixelY Then shpExtensionsSel.Top = K * lblExtensionsMnu(1).Height + 1 * Screen.TwipsPerPixelY
        shpExtensionsSel.Tag = K + 1
        If Not shpExtensionsSel.Visible = True Then shpExtensionsSel.Visible = True
    Else
        If Not shpExtensionsSel.Visible = False Then shpExtensionsSel.Visible = False
        shpExtensionsSel.Tag = ""
    End If
Else
    If Not shpExtensionsSel.Visible = False Then shpExtensionsSel.Visible = False
    shpExtensionsSel.Tag = ""
End If

End Sub

Private Sub picExtensionsMnu_Resize()
shpExtensionsMnu.Width = picExtensionsMnu.Width
shpExtensionsMnu.Height = picExtensionsMnu.Height

End Sub


Private Sub picFile_Resize()
txtFile.Top = (picFile.Height - txtFile.Height) / 2 + 1
txtFile.Left = 2
txtFile.Width = picFile.Width - 4

picFileBack.BackColor = txtFile.BackColor
picFileBack.Top = 1
picFileBack.Left = 1
picFileBack.Width = picFile.Width - 2
picFileBack.Height = picFile.Height - 2

End Sub

Private Sub PicMnuLeft_Resize()
'DRAWS A FRAME AROUND THE LEFT MENU
PicMnuLeft.Cls
PicMnuLeft.Line (0, 0)-(PicMnuLeft.Width, 0)
PicMnuLeft.Line (0, PicMnuLeft.Height - 1)-(PicMnuLeft.Width, PicMnuLeft.Height - 1)
PicMnuLeft.Line (0, 0)-(0, PicMnuLeft.Height)
PicMnuLeft.Line (PicMnuLeft.Width - 1, 0)-(PicMnuLeft.Width - 1, PicMnuLeft.Height)

'SETS LAYOUT
If picMnuLeftBackground.Height > PicMnuLeft.Height - 2 Then
    btnDOWN.Top = PicMnuLeft.Height - btnDOWN.Height
    If btnUP.Visible = False And picMnuLeftBackground.Top = 1 Then
        picMnuLeftBackground.Top = btnUP.Height
    End If
    picMnuLeftBackground.ZOrder 1
    
    btnUP.Visible = True
    btnDOWN.Visible = True
    
    If picMnuLeftBackground.Top + picMnuLeftBackground.Height < btnDOWN.Top Then
        picMnuLeftBackground.Top = btnDOWN.Top - picMnuLeftBackground.Height
    End If
Else
    picMnuLeftBackground.Top = 1
    btnUP.Visible = False
    btnDOWN.Visible = False
    
End If
PicMnuLeft.Refresh



End Sub


Private Sub PicMnuViews_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Integer
If X < 0 Or Y < 0 Or X > PicMnuViews.Width Or Y > PicMnuViews.Height Then

Else
    For ICNT = 1 To btnViewsMenu.Count - 1
        If Button = vbLeftButton And X >= 0 And X < PicMnuViews.Width And Y >= btnViewsMenu(ICNT).Top And Y < btnViewsMenu(ICNT).Top + btnViewsMenu(ICNT).Height Then
            FileView.View = ICNT - 1
            DialogView = FileView.View
            mnuIcon(4).Picture = MnuViewIcon(ICNT - 1).Picture
            DrawBtnMnu iBtnViewsIndex, mnuBtn(iBtnViewsIndex).sTag
        End If
    Next ICNT
End If

ReleaseCapture
PicMnuViews.Visible = False
picShdwH.Visible = False
picShdwV.Visible = False

DrawBtnMnu iBtnViewsIndex, mnuBtn(iBtnViewsIndex).sTag
mnuBtn(iBtnViewsIndex).UserControl_MouseOut vbLeftButton, 0, -1, -1

End Sub

Private Sub PicMnuViews_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Integer

For ICNT = 1 To btnViewsMenu.Count - 1
    If X >= 0 And X < PicMnuViews.Width And Y >= btnViewsMenu(ICNT).Top And Y < btnViewsMenu(ICNT).Top + btnViewsMenu(ICNT).Height Then
        btnViewsMenu(ICNT).SetState "H"
    Else
        btnViewsMenu(ICNT).SetState "N"
    End If

Next ICNT
End Sub


Private Sub PicMnuViews_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PicMnuViews.Visible = True Then
    Dim Rec As RECT
    Dim Rec2 As RECT
    GetWindowRect mnuBtn(iBtnViewsIndex).hwnd, Rec
    GetWindowRect PicMnuViews.hwnd, Rec2
    Dim ix As Integer
    Dim iy As Integer
    
    ix = Rec2.Left - Rec.Left + X
    iy = Rec2.Top - Rec.Top + Y
    
    If ix < 0 Or iy < 0 Or iy > mnuBtn(iBtnViewsIndex).Height Or ix > mnuBtn(iBtnViewsIndex).Width Then
        PicMnuViews_MouseDown 0, 0, -1, -1
    End If
End If


End Sub

Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picRight.MousePointer <> 0 Then Resizing = True
MouseX = X

End Sub

Private Sub picRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.iSizable = False Then Exit Sub

If Button = vbLeftButton Then
    If Resizing = True Then
        If frmDialog.Width <> frmDialog.Width + (X - MouseX) * Screen.TwipsPerPixelX Then
            If frmDialog.Width + (X - MouseX) * Screen.TwipsPerPixelX < MinWidth Then
                frmDialog.Move frmDialog.Left, frmDialog.Top, MinWidth
            Else
                frmDialog.Move frmDialog.Left, frmDialog.Top, frmDialog.Width + (X - MouseX) * Screen.TwipsPerPixelX
            End If
        
            frmDialog.Refresh
        End If
    End If
Else
    If X >= picRight.Width - 3 Then
        picRight.MousePointer = 9
    Else
        picRight.MousePointer = 0
    End If
End If

End Sub

Private Sub picRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Resizing = False

End Sub

Private Sub SetLeftToolbar()
Dim ICNT As Integer
On Error Resume Next
'Standart arrays
For ICNT = 1 To btn.Count - 1
    Unload btn(ICNT)
Next ICNT


SetLeftToolbarBtn DirView.DesktopFolder
btn(btn.Count - 1).Top = 0

SetLeftToolbarBtn DirView.MyDocumentsFolder
On Error Resume Next
btn(btn.Count - 1).Top = btn(btn.Count - 2).Top + btn(btn.Count - 2).Height

SetLeftToolbarBtn "#MYCOMPUTER"
btn(btn.Count - 1).Top = btn(btn.Count - 2).Top + btn(btn.Count - 2).Height

If FileView.MyMusicPath <> "" Then
    SetLeftToolbarBtn FileView.MyMusicPath
    btn(btn.Count - 1).Top = btn(btn.Count - 2).Top + btn(btn.Count - 2).Height
End If

If FileView.MyVideosPath <> "" Then
    SetLeftToolbarBtn FileView.MyVideosPath
    btn(btn.Count - 1).Top = btn(btn.Count - 2).Top + btn(btn.Count - 2).Height
End If

'Additional arrays
ArrayButtons = 0
ArrayButtonPath(1) = "C:\"
ArrayButtonPath(2) = "D:\"


If ArrayButtons > 0 Then
    For ICNT = 1 To ArrayButtons
        SetLeftToolbarBtn ArrayButtonPath(ICNT)
        btn(btn.Count - 1).Top = btn(btn.Count - 2).Top + btn(btn.Count - 2).Height
    Next ICNT
End If

picMnuLeftBackground.Height = btn(btn.Count - 1).Top + btn(btn.Count - 1).Height
picMnuLeftBackground.Width = btn(1).Width
picMnuLeftBackground.Left = 1
PicMnuLeft.Width = picMnuLeftBackground.Width + 2
picMnuLeftBackground.Top = 1
PicMenuLeft.Width = PicMnuLeft.Width + 8
PicMnuLeft.Left = 4
PicMnuLeft.BackColor = picToolbar.BackColor
PicMnuLeft.ForeColor = AlphaBlend(picToolbar.BackColor, vb3DDKShadow, 70)



'dRAWS BUTTON UP AND DOWN IMAGES
    'BTNUP
    picBuff.Cls
    picBuff.BackColor = PicMenuLeft.BackColor 'AlphaBlend(picToolbar.BackColor, vbWhite, 150)

    picBuff.Width = PicMnuLeft.Width
    picBuff.Height = 22
    
    picBuff.ForeColor = PicMnuLeft.ForeColor
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)

    picBuff3.AutoSize = False
    picBuff3.Width = UPImage.Width
    picBuff3.Height = UPImage.Height
    
    picBuff3.Picture = LoadPicture("")
    picBuff3.Cls
    
    picBuff3.BackColor = picBuff.BackColor
    
    picBuff3.Picture = UPImage.Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set btnUP.NormalImage = picBuff.Image
    
    picBuff3.Cls
    
    
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
        
    picBuff.Line (1, 1)-(picBuff.Width, 1)
    picBuff.Line (1, 1)-(1, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 2, 1)-(picBuff.Width - 2, picBuff.Height - 1)
    picBuff.Line (1, picBuff.Height - 2)-(picBuff.Width, picBuff.Height - 2)
    
    picBuff.ForeColor = PicMnuLeft.ForeColor
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    
    picBuff.ForeColor = PicMenuLeft.BackColor
    picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
    
    picBuff3.BackColor = picBuff.BackColor
    
    picBuff3.Picture = UPImage.Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set btnUP.FocusedImage = picBuff.Image
    
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)
    
    picBuff.Line (1, 1)-(picBuff.Width, 1)
    picBuff.Line (1, 1)-(1, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 2, 1)-(picBuff.Width - 2, picBuff.Height - 1)
    picBuff.Line (1, picBuff.Height - 2)-(picBuff.Width, picBuff.Height - 2)
    
    picBuff.ForeColor = PicMnuLeft.ForeColor
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    
    picBuff.ForeColor = PicMenuLeft.BackColor
    picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
    
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.Picture = UPImage.Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set btnUP.PressedImage = picBuff.Image

    'BTNDOWN
    picBuff.Cls
    picBuff.BackColor = PicMenuLeft.BackColor 'AlphaBlend(picToolbar.BackColor, vbWhite, 150)

    picBuff.Width = PicMnuLeft.Width
    picBuff.Height = 22
    
    picBuff.ForeColor = PicMnuLeft.ForeColor
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)

    picBuff3.AutoSize = False
    picBuff3.Width = UPImage.Width
    picBuff3.Height = UPImage.Height
    
    picBuff3.Picture = LoadPicture("")
    picBuff3.Cls
    
    picBuff3.BackColor = picBuff.BackColor
    
    picBuff3.Picture = DOWNImage.Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set btnDOWN.NormalImage = picBuff.Image
    
    picBuff3.Cls
    
    
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
        
    picBuff.Line (1, 1)-(picBuff.Width, 1)
    picBuff.Line (1, 1)-(1, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 2, 1)-(picBuff.Width - 2, picBuff.Height - 1)
    picBuff.Line (1, picBuff.Height - 2)-(picBuff.Width, picBuff.Height - 2)
    
    picBuff.ForeColor = PicMnuLeft.ForeColor
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    
    picBuff.ForeColor = PicMenuLeft.BackColor
    picBuff.Line (1, 0)-(picBuff.Width - 1, 0)
    
    picBuff3.BackColor = picBuff.BackColor
    
    picBuff3.Picture = DOWNImage.Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set btnDOWN.FocusedImage = picBuff.Image
    
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)
    
    picBuff.Line (1, 1)-(picBuff.Width, 1)
    picBuff.Line (1, 1)-(1, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 2, 1)-(picBuff.Width - 2, picBuff.Height - 1)
    picBuff.Line (1, picBuff.Height - 2)-(picBuff.Width, picBuff.Height - 2)
    
    picBuff.ForeColor = PicMnuLeft.ForeColor
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    
    picBuff.ForeColor = PicMenuLeft.BackColor
    picBuff.Line (1, 0)-(picBuff.Width - 1, 0)
    
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.Picture = DOWNImage.Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    picBuff.Refresh
    
    Set btnDOWN.PressedImage = picBuff.Image





PicMnuLeft_Resize
Form_Resize

End Sub

Private Sub SetLeftToolbarBtn(Tag As String)
Dim index As Integer
index = btn.Count

picBuff.Width = 85
picBuff.Height = 67

Dim Title As String
Dim Title1 As String
Dim Title2 As String
Dim iPos As String
Dim K As Integer

Dim title2Len As Integer

If Tag = DirView.MyDocumentsFolder Then
    Title = DirView.MyDocuments
ElseIf Tag <> "#MYCOMPUTER" Then
    If Len(Tag) > 3 Then
        Title = Mid(Tag, InStrRev(Tag, "\") + 1)
    Else
        Title = Tag
    End If
Else
    Title = DirView.MyComputer
End If

iPos = 0

Title1 = Title

If picBuff3.TextWidth(Title1) > picBuff.Width - 4 Then
    For K = 1 To Len(Title1)
        iPos = InStrRev(Title1, " ")
        
        If iPos > 0 Then Title1 = Left(Title1, iPos - 1)
        
        If picBuff3.TextWidth(Title1) <= picBuff.Width - 4 Then
            Exit For
        End If
    Next K
    
    
    If iPos = 0 Then
        Title1 = Title
        For K = 1 To Len(Title)
            Title1 = Left(Title1, Len(Title1) - 1)
            If picBuff3.TextWidth(Title1 & "...") <= picBuff.Width - 4 Then
                Exit For
            End If
        Next K
        Title1 = Title1 & "..."
    Else
        Title2 = Mid(Title, Len(Title1) + 2)
        
        If picBuff3.TextWidth(Title2) > picBuff.Width - 4 Then
            title2Len = Len(Title2)
            For K = 1 To title2Len
                Title2 = Left(Title2, Len(Title2) - 1)
                If picBuff3.TextWidth(Title2 & "...") <= picBuff.Width - 4 Then
                    Exit For
                End If
            Next K
            Title2 = Title2 & "..."
        End If
    End If
End If
On Error Resume Next
Load btn(index)
btn(index).SetTag Tag

picBuff.Cls
picBuff.BackColor = picToolbar.BackColor

picBuff3.AutoSize = False
picBuff3.Width = 32
picBuff3.Height = 32

picBuff3.Picture = LoadPicture("")
picBuff3.Cls

picBuff3.BackColor = picBuff.BackColor

Dim hLargeIcon As Long
Dim hSmallIcon As Long
                
If Tag = "#MYCOMPUTER" Then
    ExtractIconEx "explorer.exe", 0, hLargeIcon, hSmallIcon, 1
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon hSmallIcon
ElseIf Tag = DirView.MyDocumentsFolder Then
    ExtractIconEx "mydocs.dll", 0, hLargeIcon, hSmallIcon, 1
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon hSmallIcon
ElseIf Tag = DirView.DesktopFolder Then
    ExtractIconEx "shell32.dll", 34, hLargeIcon, hSmallIcon, 1
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon hSmallIcon
Else
    iIcon = SHGetFileInfo(Tag, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    ImageList_Draw iIcon, SHInfo.iIcon, picBuff3.hDC, 0, 0, ILD_TRANSPARENT
End If

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.Width) / 2), 3, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY


picBuff3.Cls

picBuff3.Height = picBuff3.TextHeight(Title1)
picBuff3.Width = picBuff3.TextWidth(Title1)

picBuff3.Font.Size = 8

picBuff3.Print Title1

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.TextWidth(Title1)) / 2), 37, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

picBuff3.Cls

picBuff3.Height = picBuff3.TextHeight(Title2)
picBuff3.Width = picBuff3.TextWidth(Title2)

picBuff3.Font.Size = 8

picBuff3.Print Title2

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.TextWidth(Title2)) / 2), 37 + picBuff3.Height, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

picBuff.Refresh

Set btn(index).NormalImage = picBuff.Image

picBuff3.Width = 32
picBuff3.Height = 32

picBuff3.Cls
picBuff.ForeColor = vbHighlight
picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)

picBuff.Line (0, 0)-(picBuff.Width, 0)
picBuff.Line (0, 0)-(0, picBuff.Height)
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)

picBuff3.BackColor = picBuff.BackColor


If Tag = "#MYCOMPUTER" Then
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
ElseIf Tag = DirView.MyDocumentsFolder Then
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
ElseIf Tag = DirView.DesktopFolder Then
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
Else
    iIcon = SHGetFileInfo(Tag, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    ImageList_Draw iIcon, SHInfo.iIcon, picBuff3.hDC, 0, 0, ILD_TRANSPARENT
End If

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.Width) / 2), 3, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
picBuff.Refresh

picBuff3.Cls

picBuff3.Height = picBuff3.TextHeight(Title1)
picBuff3.Width = picBuff3.TextWidth(Title1)

picBuff3.Font.Size = 8

picBuff3.Print Title1

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.TextWidth(Title1)) / 2), 37, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

picBuff3.Cls

picBuff3.Height = picBuff3.TextHeight(Title2)
picBuff3.Width = picBuff3.TextWidth(Title2)

picBuff3.Font.Size = 8

picBuff3.Print Title2

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.TextWidth(Title2)) / 2), 37 + picBuff3.Height, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

picBuff3.Cls

Set btn(index).FocusedImage = picBuff.Image

picBuff3.Width = 32
picBuff3.Height = 32

picBuff.ForeColor = vbHighlight
picBuff.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

picBuff.Line (0, 0)-(picBuff.Width, 0)
picBuff.Line (0, 0)-(0, picBuff.Height)
picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)

picBuff3.BackColor = picBuff.BackColor

If Tag = "#MYCOMPUTER" Then
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon hLargeIcon
ElseIf Tag = DirView.MyDocumentsFolder Then
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon hLargeIcon
ElseIf Tag = DirView.DesktopFolder Then
    DrawIconEx picBuff3.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon hLargeIcon
Else
    iIcon = SHGetFileInfo(Tag, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    ImageList_Draw iIcon, SHInfo.iIcon, picBuff3.hDC, 0, 0, ILD_TRANSPARENT
End If
BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.Width) / 2) + 1, 4, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
picBuff.Refresh

picBuff3.Cls

picBuff3.Height = picBuff3.TextHeight(Title1)
picBuff3.Width = picBuff3.TextWidth(Title1)

picBuff3.Font.Size = 8

picBuff3.Print Title1

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.TextWidth(Title1)) / 2) + 1, 38, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

picBuff3.Cls

picBuff3.Height = picBuff3.TextHeight(Title2)
picBuff3.Width = picBuff3.TextWidth(Title2)

picBuff3.Font.Size = 8

picBuff3.Print Title2

BitBlt picBuff.hDC, Int((picBuff.Width - picBuff3.TextWidth(Title2)) / 2) + 1, 38 + picBuff3.Height, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

picBuff3.Cls

Set btn(index).PressedImage = picBuff.Image

btn(index).ZOrder
btn(index).Visible = True

picBuff3.Font.Size = 9

End Sub

Private Sub Picture1_Paint()
Dim c         As ColorAndAlpha
OleTranslateColor Me.Picture1.BackColor, 0, VarPtr(c)

If RGB(c.r, c.G, c.B) <> ColorBack Then 'user changed desktop theam - so we need to redraw all the objects...
    ColorBack = RGB(c.r, c.G, c.B)
    picToolbar.BackColor = Me.BackColor
    RedrawToolBar
    SetLeftToolbar
    Me.picResize.BackColor = Me.BackColor
    DirView.UserControl_Initialize
    FileView.Refresh
    RedrawMnuExtensions
End If


End Sub

Private Sub tmrLeftMenu_Timer()
Dim cc As Integer
cc = 8
If tmrLeftMenu.Interval > 50 Then
    tmrLeftMenu.Interval = 50
End If

If tmrLeftMenu.Tag = "UP" Then
    If picMnuLeftBackground.Top < btnUP.Height - cc Then
        picMnuLeftBackground.Top = picMnuLeftBackground.Top + cc
    Else
        picMnuLeftBackground.Top = btnUP.Height
        tmrLeftMenu.Enabled = False
    End If
ElseIf tmrLeftMenu.Tag = "DOWN" Then
    If picMnuLeftBackground.Top > btnDOWN.Top - picMnuLeftBackground.Height + cc Then
        picMnuLeftBackground.Top = picMnuLeftBackground.Top - cc
    Else
        picMnuLeftBackground.Top = btnDOWN.Top - picMnuLeftBackground.Height
        tmrLeftMenu.Enabled = False
    End If
End If

End Sub

Private Sub SetMenuViews()
Dim ICNT As Integer
Dim cc As cMemDC

'SETS BUTTONS
For ICNT = 1 To Me.btnViewsMenu.Count - 1
    Unload Me.btnViewsMenu(ICNT)
Next ICNT

Dim tWidth As Integer
tWidth = picBuff3.TextWidth(iView0)
If tWidth < picBuff3.TextWidth(iView1) Then tWidth = picBuff3.TextWidth(iView1)
If tWidth < picBuff3.TextWidth(iView2) Then tWidth = picBuff3.TextWidth(iView2)
If tWidth < picBuff3.TextWidth(iView3) Then tWidth = picBuff3.TextWidth(iView3)
If tWidth < picBuff3.TextWidth(iView4) Then tWidth = picBuff3.TextWidth(iView4)

For ICNT = 1 To 4
    Load Me.btnViewsMenu(ICNT)


    picBuff.Height = 24
    
    picBuff.Width = 24 + 5 + tWidth
    picBuff.BackColor = vbWhite
    
    picBuff3.Picture = LoadPicture("")
    picBuff3.Cls
    
    picBuff3.Width = 24
    picBuff3.Height = 24
    
    Dim firsColor As OLE_COLOR
    Dim SecondColor As OLE_COLOR
    
    firsColor = vb3DHighlight
    SecondColor = AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240)

    Dim IcNT2 As Integer
    picBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 40)
    For IcNT2 = 1 To Me.picBuff3.Width
        picBuff3.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * IcNT2 / picBuff3.Width))
        picBuff3.Line (IcNT2 - 1, 0)-(IcNT2 - 1, picBuff3.Height)
        picBuff3.Refresh
    Next IcNT2
    
    If FileView.View = ICNT - 1 Then
        picBuff3.ForeColor = AlphaBlend(vbHighlight, vbWindowBackground, 150)
        '
        picBuff3.Line (1, 1)-(picBuff3.Width - 1, 1)
        picBuff3.Line (1, picBuff3.Height - 2)-(picBuff3.Width - 1, picBuff3.Height - 2)
        picBuff3.Line (1, 1)-(1, picBuff3.Height - 1)
        picBuff3.Line (picBuff3.Width - 2, 1)-(picBuff3.Width - 2, picBuff3.Height - 1)
    
        picBuff3.ForeColor = picBuff3.BackColor
        picBuff3.Line (2, 2)-(picBuff3.Width - 2, 2)
        picBuff3.Line (2, 3)-(picBuff3.Width - 2, 3)
        
        picBuff3.Line (2, picBuff3.Height - 3)-(picBuff3.Width - 2, picBuff3.Height - 3)
        picBuff3.Line (2, picBuff3.Height - 4)-(picBuff3.Width - 2, picBuff3.Height - 4)
        
        
        picBuff3.Line (2, 2)-(2, picBuff3.Height - 2)
        picBuff3.Line (3, 2)-(3, picBuff3.Height - 2)
        
        picBuff3.Line (picBuff3.Width - 3, 2)-(picBuff3.Width - 3, picBuff3.Height - 2)
        picBuff3.Line (picBuff3.Width - 4, 2)-(picBuff3.Width - 4, picBuff3.Height - 2)

    
    Else
        'picBuff3.BackColor = picToolbar.BackColor
        
    End If
    
    BitBlt picBuff.hDC, 0, 0, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    
    picBuff3.Cls
    picBuff3.Picture = MnuViewIcon(ICNT - 1)
    
    If FileView.View <> ICNT - 1 Then
        Set cc = DimBitmap2(picBuff3.Picture, picBuff3.BackColor, False)
        cc.BitBlt picBuff.hDC, (picBuff.Height - picBuff3.Height) / 2, (picBuff.Height - picBuff3.Height) / 2, 16, 16, 0, 0
    Else
        Set cc = DimBitmap2(picBuff3.Picture, picBuff3.BackColor, True)
        cc.BitBlt picBuff.hDC, (picBuff.Height - picBuff3.Height) / 2, (picBuff.Height - picBuff3.Height) / 2, 16, 16, 0, 0
    End If
    
    picBuff3.Cls
    picBuff3.Picture = LoadPicture("")
    picBuff3.Width = tWidth
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.ForeColor = &H80000007
    
    If ICNT = 1 Then
        picBuff3.Print iView0
    ElseIf ICNT = 2 Then
        picBuff3.Print iView1
    ElseIf ICNT = 3 Then
        picBuff3.Print iView2
    ElseIf ICNT = 4 Then
        picBuff3.Print iView3
    ElseIf ICNT = 5 Then
        picBuff3.Print iView4
    End If
    
    BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    
    Set btnViewsMenu(ICNT).NormalImage = picBuff.Image
    
    picBuff.Cls
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
    
    If FileView.View = ICNT - 1 Then
        picBuff3.Width = 24
        picBuff3.Height = 24
        picBuff3.ForeColor = AlphaBlend(vbHighlight, vbWindowBackground, 150)
        picBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
        
        picBuff3.Line (1, 1)-(picBuff3.Width - 1, 1)
        picBuff3.Line (1, picBuff3.Height - 2)-(picBuff3.Width - 1, picBuff3.Height - 2)
        picBuff3.Line (1, 1)-(1, picBuff3.Height - 1)
        picBuff3.Line (picBuff3.Width - 2, 1)-(picBuff3.Width - 2, picBuff3.Height - 1)
        BitBlt picBuff.hDC, 0, 0, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    Else
         picBuff3.BackColor = picBuff.BackColor
    End If
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)

    picBuff3.Picture = MnuViewIcon(ICNT - 1)
    
    'Set cc = GetHoverImage(picBuff3.Picture, vbWhite)
    'BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

    BitBlt picBuff.hDC, (picBuff.Height - picBuff3.Width) / 2, (picBuff.Height - picBuff3.Height) / 2, 18, 18, picBuff3.hDC, 0, 0, SRCCOPY

    picBuff3.Cls
    picBuff3.Picture = LoadPicture("")
    picBuff3.Width = tWidth
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.ForeColor = &H80000007
    
    If ICNT = 1 Then
        picBuff3.Print iView0
    ElseIf ICNT = 2 Then
        picBuff3.Print iView1
    ElseIf ICNT = 3 Then
        picBuff3.Print iView2
    ElseIf ICNT = 4 Then
        picBuff3.Print iView3
    ElseIf ICNT = 5 Then
        picBuff3.Print iView4
    End If
    
    BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picBuff3.Height) / 2, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY

    Set btnViewsMenu(ICNT).FocusedImage = picBuff.Image
    Set btnViewsMenu(ICNT).PressedImage = picBuff.Image
    
    btnViewsMenu(ICNT).Left = 2
    btnViewsMenu(ICNT).Top = 2 + (ICNT - 1) * btnViewsMenu(ICNT).Height
    btnViewsMenu(ICNT).Visible = True
Next ICNT
'SETS THE WHOLE MENU

PicMnuViews.Width = picBuff.Width + 4
PicMnuViews.Height = picBuff.Height * (btnViewsMenu.Count - 1) + 4

shpBorder.BorderColor = AlphaBlend(picToolbar.BackColor, vb3DDKShadow, 150)
shpBorder.Left = 0
shpBorder.Top = 0
shpBorder.Width = PicMnuViews.Width
shpBorder.Height = PicMnuViews.Height

lnMnuView.X1 = PicMnuViews.Width - mnuBtn(iBtnViewsIndex).Width + 1
lnMnuView.X2 = PicMnuViews.Width - 1
lnMnuView.Y1 = 0
lnMnuView.Y2 = 0

PicMnuViews.Top = picToolbar.Top + mnuBtn(iBtnViewsIndex).Top + mnuBtn(iBtnViewsIndex).Height - 1
PicMnuViews.Left = picToolbar.Left + mnuBtn(iBtnViewsIndex).Left + mnuBtn(iBtnViewsIndex).Width - PicMnuViews.Width

'DRAWS SHADDOWS ARROUND THE MENU...
picShdwH.Top = PicMnuViews.Top + PicMnuViews.Height
picShdwH.Left = PicMnuViews.Left + 5
picShdwH.Width = PicMnuViews.Width - 4
picShdwH.Height = 4


picShdwV.Left = PicMnuViews.Left + PicMnuViews.Width
picShdwV.Top = PicMnuViews.Top - mnuBtn(iBtnViewsIndex).Height + 5
picShdwV.Height = PicMnuViews.Height + mnuBtn(iBtnViewsIndex).Height - 1
picShdwV.Width = 4

picShdwV.ZOrder
picShdwH.ZOrder
Dim plusY As Integer 'pixels from top of the form to top of the form surface
Dim plusX As Integer

plusX = (frmDialog.Width / Screen.TwipsPerPixelX - frmDialog.ScaleWidth) / 2
plusY = (frmDialog.Height / Screen.TwipsPerPixelY - frmDialog.ScaleHeight - plusX)

DeskHdc = GetDC(0)
ret = BitBlt(picShdwH.hDC, 0, 0, picShdwH.Width, picShdwH.Height, DeskHdc, frmDialog.Left / Screen.TwipsPerPixelX + picShdwH.Left + plusX, frmDialog.Top / Screen.TwipsPerPixelY + picShdwH.Top + plusY, SRCCOPY)
ret = ReleaseDC(0&, DeskHdc)
    
picShdwH.Picture = picShdwH.Image

Set cc = DrawShadow(picShdwH.Picture, False)
cc.BitBlt picShdwH.hDC, 0, 0, picShdwH.Width, picShdwH.Height, 0, 0

DeskHdc = GetDC(0)
ret = BitBlt(picShdwV.hDC, 0, 0, picShdwV.Width, picShdwV.Height, DeskHdc, frmDialog.Left / Screen.TwipsPerPixelX + picShdwV.Left + plusX, frmDialog.Top / Screen.TwipsPerPixelY + picShdwV.Top + plusY, SRCCOPY)
ret = ReleaseDC(0&, DeskHdc)

picShdwV.Picture = picShdwV.Image

Set cc = DrawShadow(picShdwV.Picture, True)
cc.BitBlt picShdwV.hDC, 0, 0, picShdwV.Width, picShdwV.Height, 0, 0

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

