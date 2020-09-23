VERSION 5.00
Begin VB.UserControl FileView 
   Alignable       =   -1  'True
   BackColor       =   &H005A371B&
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   5040
      Top             =   1200
   End
   Begin VB.PictureBox PicMnu 
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
      Left            =   2760
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      Begin VB.PictureBox picBan 
         AutoRedraw      =   -1  'True
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
         Height          =   135
         Index           =   0
         Left            =   120
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin OricronDialog.Button btnMenu 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         FillColor       =   &H8000000E&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   120
         Top             =   -240
         Width           =   1575
      End
   End
   Begin VB.PictureBox PicMnuExtendet 
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
      Left            =   2040
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      Begin VB.PictureBox picBanExtendet 
         AutoRedraw      =   -1  'True
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
         Height          =   135
         Index           =   0
         Left            =   120
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin OricronDialog.Button btnMenuExtendet 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.Line Line1 
         X1              =   8
         X2              =   16
         Y1              =   16
         Y2              =   40
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
   Begin VB.Timer tmrScr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   1680
   End
   Begin VB.PictureBox picScrH 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
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
      Height          =   255
      Left            =   2040
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   16
      Top             =   4560
      Width           =   3495
      Begin VB.PictureBox picScrollerH 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   19
         Top             =   0
         Width           =   1815
         Begin VB.Line lnH 
            X1              =   72
            X2              =   72
            Y1              =   0
            Y2              =   16
         End
      End
      Begin OricronDialog.Button btnRIGHT 
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   450
      End
      Begin OricronDialog.Button btnLEFT 
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   450
      End
   End
   Begin VB.PictureBox picMnuText 
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
      Left            =   2880
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicBuff3 
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
      Left            =   1440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picHideDetails 
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
      Height          =   135
      Left            =   4680
      Picture         =   "FileView.ctx":0000
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picExtendet 
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
      Left            =   2880
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2160
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      BackColor       =   &H8000000E&
      Height          =   450
      Left            =   2400
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picScr 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
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
      Height          =   4890
      Left            =   5760
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   6
      Top             =   0
      Width           =   255
      Begin OricronDialog.Button btnDOWN 
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   4200
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   450
      End
      Begin OricronDialog.Button btnUP 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   450
      End
      Begin VB.PictureBox picScroller 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   0
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   9
         Top             =   1080
         Width           =   255
         Begin VB.Line LN 
            X1              =   0
            X2              =   16
            Y1              =   112
            Y2              =   120
         End
      End
   End
   Begin VB.PictureBox picDetails 
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
      Height          =   135
      Left            =   4680
      Picture         =   "FileView.ctx":04FE
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox IconBuff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000015&
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
      Height          =   105
      Left            =   4320
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picBackground 
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
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   4
      Top             =   3240
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
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picOzadje 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox picShdw 
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
         Left            =   1080
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox PicRename 
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
         Left            =   600
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
         Begin VB.TextBox txtRename 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox lstExtensions 
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin OricronDialog.Button btn 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
      End
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
      Left            =   1440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   0
      Left            =   4680
      Picture         =   "FileView.ctx":09FC
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   1
      Left            =   4680
      Picture         =   "FileView.ctx":0F86
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "FileView.ctx":1510
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuViewIcon 
      Height          =   240
      Index           =   3
      Left            =   4680
      Picture         =   "FileView.ctx":1A9A
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuIcon 
      Height          =   240
      Index           =   3
      Left            =   4320
      Picture         =   "FileView.ctx":2024
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuIcon 
      Height          =   240
      Index           =   2
      Left            =   4320
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuIcon 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "FileView.ctx":25AE
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MnuIconExtendet 
      Height          =   240
      Index           =   0
      Left            =   4680
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image MnuIcon 
      Height          =   240
      Index           =   0
      Left            =   4320
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Icon 
      Height          =   255
      Index           =   0
      Left            =   4800
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image IconS 
      Height          =   255
      Index           =   0
      Left            =   4800
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   3
      Left            =   4320
      Picture         =   "FileView.ctx":2B38
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   2
      Left            =   4320
      Picture         =   "FileView.ctx":3012
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   1
      Left            =   4320
      Picture         =   "FileView.ctx":34EC
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   0
      Left            =   4320
      Picture         =   "FileView.ctx":39C6
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image LIconS 
      Height          =   255
      Index           =   0
      Left            =   4320
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image LIcon 
      Height          =   255
      Index           =   0
      Left            =   4320
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FileView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim sqy As Boolean

Dim DeskHdc&, ret&
Dim mx As Single
Dim my As Single
Dim KCY As Boolean
Dim ColumnWidth As Integer
    
Public MyMusicPath As String
Public MyVideosPath As String

    
Dim dPath As String
Dim DoGen As Boolean
Dim qk As Boolean
Private Type Whatever
    name As String
    Extension As String
    FullName As String
    Selected As Boolean
    ShowDetails As Boolean
    Details As String
    IconIndex As Long
    selIcon As PictureBox
End Type

'menu captions
Public MNU_SELECT As String
Public MNU_VIEW As String
Public MNU_COPY As String
Public MNU_CUT As String
Public MNU_PASTE As String
Public MNU_RENAME As String
Public MNU_NEWFOLDER As String
Public MNU_PROPERTYS As String
Public MNU_REFRESH As String
Public MNU_DELETE As String

'extendet menu captions
Public EXTMNU_EXPANDALL As String
Public EXTMNU_UNEXPANDALL As String
Public EXTMNU_EXPAND As String
Public EXTMNU_UNEXPAND As String

'other captions
Public CAPTION_LASTACCESSED As String
Public CAPTION_FILESIZE As String
Public CAPTION_FILETYPE As String
Public CAPTION_FOLDERPATH As String
Public CAPTION_FOLDERSIZE As String
Public CAPTION_FILESINFOLDER As String
Public CAPTION_DRIVEFREESPACE  As String
Public CAPTION_DRIVESIZE  As String
Public CAPTION_FILESYSTEM  As String
Public CAPTION_DEVICEUNAVAILABLE  As String

Public MNUVIEW_ETENDETMODE  As String
Public MNUVIEW_LIST_SMALL  As String
Public MNUVIEW_LIST_LARGE  As String
Public MNUVIEW_ICONS  As String

Dim qPathExists As Boolean
Dim iFile() As Whatever

Public MultiSelect As Boolean
Public DesktopFolder As String

Dim iTop As Long
Dim iLeft As Integer
Dim iHeight As Long
Dim SelHeight As Integer
Dim selDetHeight As Integer
Public DoNotGenerate As Boolean
Dim MouseY As Single
Dim MouseX As Single
Dim bSelected As Long
Dim sSelected As Long

Dim FSO As FileSystemObject

Public ListCount As Long
Dim RowsInColumnCount As Integer    'how many lines can go into one column

'View Types
Dim dView As Integer
Dim dScrollersType As Integer

Dim ColumnCount As Integer        'Numbers of columns in listview
    
'Events
Public Event ErrorOcured(Error As Long)
Public Event DirSelect(Path As String)
Public Event PathChange(NewPath As String, OldPath As String)
Public Event FileSelect(Files As ListBox)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event FilePreSelect(Files As ListBox)
Public Event MenuNewFolderClick()
Public Event ViewChange(NewView As Integer)
Public Event DeletableItemSelected(AnyDeletableItemSelected As Boolean)


'Colors
Public SelBorderColor As OLE_COLOR
Public SelBackColor As OLE_COLOR
Public SelForeColor As OLE_COLOR
Public BackColor As OLE_COLOR
Public ForeColor As OLE_COLOR

'FUNCTIONS:
'Graphics
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Dim iIcon As Long           'icon of the folder we get from shell

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

'To get icons from files...
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long

'This code is written by Tim Misset - to get icons from folders i took some code from him - thanks;)
'Const LARGE_ICON As Integer = 32                '  do not need that so i removed it
'Const SMALL_ICON As Integer = 16                '  do not need that so i removed it
Const MAX_PATH = 260
Const ILD_TRANSPARENT = &H1                      '  Display transparent
Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Const SHGFI_EXETYPE = &H2000                     '  return exe type
Const SHGFI_LARGEICON = &H0                      '  get large icon
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

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'AlphaBlending - to get blendet colors (buttons etc.)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)

Private Type ColorAndAlpha
    r                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
End Type

Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long
    lpClass       As String
    hProcess      As Long
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32" _
    (SEI As SHELLEXECUTEINFO) As Long

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Public Enum RemoveMethod
    RecycleFile = 1
    DeleteFile = 2
End Enum

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_CREATEPROGRESSDLG As Long = &H0

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Sub DeleteSelected()
Dim ICNT As Long
Dim ref As Boolean
For ICNT = 1 To ListCount
    If iFile(ICNT).Selected = True Then
        If RemoveFile(iFile(ICNT).FullName, RecycleFile) = True Then
            ref = True
        End If
    End If
Next ICNT
If ref = True Then Refresh

End Sub

Private Sub DrawScrollers()
If dScrollersType = 0 Then
    'Draws buttons - just read it line by line - you should see what it does
    picBuff.Width = 15
    picBuff.Height = 15
    picBuff.Cls
    
    picBuff.BackColor = &H8000000F
    
    picBuff.ForeColor = &H80000014
    picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height - 1)
    
    picBuff.ForeColor = &H80000010
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
    
    IconBuff.BackColor = picBuff.BackColor
    IconBuff.Picture = btnSCR(0).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnUP.NormalImage = picBuff.Image
    Set btnUP.FocusedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(1).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnDOWN.NormalImage = picBuff.Image
    Set btnDOWN.FocusedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(2).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnLEFT.NormalImage = picBuff.Image
    Set btnLEFT.FocusedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(3).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnRIGHT.NormalImage = picBuff.Image
    Set btnRIGHT.FocusedImage = picBuff.Image
    
    picBuff.ForeColor = &H80000010
    picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

    IconBuff.Picture = btnSCR(0).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnUP.PressedImage = picBuff.Image

    IconBuff.Picture = btnSCR(1).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnDOWN.PressedImage = picBuff.Image
  
    IconBuff.Picture = btnSCR(2).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnLEFT.PressedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(3).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set btnRIGHT.PressedImage = picBuff.Image
    
    'Draw scrollers - follow it line by line;p
    picScr.Width = btnUP.Width
    picScroller.Width = btnUP.Width
     
    picScroller.Height = 1700 ' this is somehow the maximum visible size on the screen - we need to set this, so it draws the whole scroller - because of bitblt
    
    picScroller.BackColor = &H8000000F
    
    picScroller.ForeColor = &H80000014
    
    picScroller.Line (0, 0)-(picScroller.Width - 1, 0)
    picScroller.Line (0, 0)-(0, picScroller.Height)
    
    picScroller.ForeColor = &H80000010
    picScroller.Line (picScroller.Width - 1, 0)-(picScroller.Width - 1, picScroller.Height)

    picScrH.Height = btnLEFT.Height
    picScrollerH.Height = btnLEFT.Height
    
    LN.X1 = 0
    LN.X2 = picScroller.Width
    LN.BorderColor = &H80000010
    
    picScrollerH.Width = 2050 ' this is somehow the maximum visible size on the screen - we need to set this, so it draws the whole scroller - because of bitblt
    
    picScrollerH.BackColor = &H8000000F
    
    picScrollerH.ForeColor = &H80000014
    
    picScrollerH.Line (0, 0)-(picScrollerH.Width, 0)
    picScrollerH.Line (0, 0)-(0, picScrollerH.Height - 1)
    
    picScrollerH.ForeColor = &H80000010
    picScrollerH.Line (0, picScrollerH.Height - 1)-(picScrollerH.Width, picScrollerH.Height - 1)

    lnH.Y1 = 0
    lnH.Y2 = picScrollerH.Height
    lnH.BorderColor = &H80000010
    
    
End If

End Sub


Private Function GetFileSize(FileName As String) As String
Dim K As Long
Dim K1

Dim c As String
K = FileLen(FileName)
K1 = K
If K > 1048576 Then
    c = K
    K = Left(K, Len(c) - 4)
    GetFileSize = Int(K * 100 / 104.8576) / 100 & " MB"
    
ElseIf K > 1024 Then
    GetFileSize = Int(K * 100 / (1024)) / 100 & " kB"
Else
    GetFileSize = K & " B"
End If

If K > 1024 Then
    GetFileSize = GetFileSize & "  (" & K1 & " B)"
End If

End Function
Private Function GetFileDate(FileName As String) As String
Dim File
Set File = FSO.GetFile(FileName)

GetFileDate = Format(File.DateLastModified, "d/ mmmm yyyy") & ", " & Format(File.DateLastModified, "hh:nn:ss")

End Function
Private Function GetFileType(FileName As String) As String
Dim File
Set File = FSO.GetFile(FileName)

GetFileType = File.Type

End Function

Private Function GetFileAtribute(FileName As String) As String
Dim File
Set File = FSO.GetFile(FileName)

GetFileAtribute = File.Attributes

End Function
Private Function GetSize(K) As String

Dim c As String
Dim K1
K1 = K
If K > 1073741824 Then
    c = K
    K = Left(K, Len(c) - 7)
    GetSize = Int(K * 100 / 107.3741824) / 100 & " GB"
ElseIf K > 1048576 Then
    c = K
    K = Left(K, Len(c) - 4)
    GetSize = Int(K * 100 / 104.8576) / 100 & " MB"
ElseIf K > 1024 Then
    GetSize = Int(K * 100 / (1024)) / 100 & " kB"
Else
    GetSize = K & " B"
End If

If sqy = True Then Exit Function
If K > 1024 Then
    GetSize = GetSize & "  (" & K1 & " B)"
End If

End Function

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

Private Function GetFileTop(index As Long) As Long
On Error Resume Next
'here we get distance of the line from the top
Dim ICNT As Integer
GetFileTop = 0
For ICNT = 1 To index - 1
    If iFile(ICNT).ShowDetails = True Then
        GetFileTop = GetFileTop + selDetHeight
    Else
        GetFileTop = GetFileTop + SelHeight
    End If
Next ICNT



End Function

Public Property Let Path(ipath As String)
If ipath = dPath Then Exit Property
bSelected = 0
Dim Bpath As String
Bpath = dPath
dPath = ipath
Dim ICNT As Long
Dim cPath As String
DoGen = True
ReDim iFile(0)
List1.Clear
Dim IcNT2 As Integer
Dim IconFound As Boolean
'On Error Resume Next
DesktopFolder = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
iLeft = 0

IconBuff.Height = 16
IconBuff.Width = 16
    
If ipath = "#MYCOMPUTER" Then
    If qk = False Then RaiseEvent PathChange(ipath, Bpath)
    RaiseEvent DirSelect(ipath)

    If DoGen = False Then Exit Property
    picOzadje.Cls
    iTop = 0
    
    List1.Clear
    
    For ICNT = 0 To Drive1.ListCount - 1
        List1.AddItem Drive1.List(ICNT)
    
    Next ICNT
    
    ListCount = List1.ListCount
    
    ReDim Preserve iFile(ListCount)
    
    For ICNT = 1 To ListCount
        iFile(ICNT).FullName = Capitalize(Left(List1.List(ICNT - 1), 2)) & "\"
        iFile(ICNT).Selected = False
        iFile(ICNT).ShowDetails = False
        
        'Here we format the drive caption to look explorer style
        If Len(Mid(List1.List(ICNT - 1), 5)) > 0 Then
            iFile(ICNT).name = Left(Mid(List1.List(ICNT - 1), 5), Len(Mid(List1.List(ICNT - 1), 5)) - 1) & " (" & Capitalize(Left(List1.List(ICNT - 1), 2)) & ")"
        Else
            iFile(ICNT).name = "(" & Capitalize(Left(List1.List(ICNT - 1), 2)) & ")"
        End If

        iFile(ICNT).Extension = "#FOLDER"

        If Icon.Count > 1 Then
            IconFound = False
            
            For IcNT2 = 1 To Icon.Count - 1 ' DRIVES HAVE THEIR OWN ICONS THEREFORE WE HAVE TO LOAD THEME EACH BY ITSSELVE
                If iFile(ICNT).FullName = Icon(IcNT2).Tag Then '
                    IconFound = True '
                    Exit For '
                End If '
            Next IcNT2 '
            
            If IconFound = True Then 'for every file we find a sutible preloaded icon... except for the fileswith .lnk xtensions (srtcuts - they each have its own icon - according to the file they point)
                iFile(ICNT).IconIndex = IcNT2
            
            Else
                Load Icon(Icon.Count)
                Load IconS(IconS.Count)
                Load LIcon(LIcon.Count)
                Load LIconS(LIconS.Count)
                
                iFile(ICNT).IconIndex = Icon.Count - 1
                Icon(Icon.Count - 1).Tag = iFile(ICNT).FullName
                
                IconBuff.Height = 16
                IconBuff.Width = 16
                
                IconBuff.Cls
                IconBuff.BackColor = BackColor
                IconBuff.Picture = LoadPicture("")
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                'iIcon = SHGetFileInfo("c:\windows\system32\shell.dll", 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

                ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                IconBuff.Refresh
                
                Icon(Icon.Count - 1).Picture = IconBuff.Image
                
                IconBuff.Cls
                IconBuff.BackColor = SelBackColor
                IconBuff.Picture = LoadPicture("")
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
            
                ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                IconBuff.Refresh
            
                IconS(Icon.Count - 1).Picture = IconBuff.Image
                
                IconBuff.Height = 32
                IconBuff.Width = 32
                
                IconBuff.Cls
                IconBuff.BackColor = BackColor
                IconBuff.Picture = LoadPicture("")
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
            
                ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                IconBuff.Refresh
            
                LIcon(Icon.Count - 1).Picture = IconBuff.Image
                
                IconBuff.Cls
                IconBuff.BackColor = SelBackColor
                IconBuff.Picture = LoadPicture("")
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
            
                ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                IconBuff.Refresh
            
                LIconS(Icon.Count - 1).Picture = IconBuff.Image
            End If
        Else
            Load Icon(Icon.Count)
            Load IconS(IconS.Count)
            Load LIcon(LIcon.Count)
            Load LIconS(LIconS.Count)
            
            iFile(ICNT).IconIndex = Icon.Count - 1
            Icon(Icon.Count - 1).Tag = iFile(ICNT).FullName
            
            IconBuff.Height = 16
            IconBuff.Width = 16
            
            IconBuff.Cls
            IconBuff.BackColor = BackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            Icon(Icon.Count - 1).Picture = IconBuff.Image
            
            IconBuff.Cls
            IconBuff.BackColor = SelBackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            IconS(Icon.Count - 1).Picture = IconBuff.Image
            
            IconBuff.Height = 32
            IconBuff.Width = 32
            
            IconBuff.Cls
            IconBuff.BackColor = BackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            LIcon(Icon.Count - 1).Picture = IconBuff.Image
            
            IconBuff.Cls
            IconBuff.BackColor = SelBackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            LIconS(Icon.Count - 1).Picture = IconBuff.Image
        End If
        
    Next ICNT
    
    'Before we Genrate it we must check the size of the list (height - to set scrollers)
    DoNotGenerate = True
    
    If dView = 0 Then
        SelHeight = 18 'NormalviewHeight
        selDetHeight = 64               '"Extendet" dView Mode
    ElseIf dView = 1 Then
        SelHeight = 18 'NormalviewHeight
        selDetHeight = 0                'NoExtendetMode here
    ElseIf dView = 2 Then
        SelHeight = 34 'NormalviewHeight
        selDetHeight = 0                'NoExtendetMode here
    End If
    
    If CheckHeight > UserControl.Height / Screen.TwipsPerPixelY Then
        btnDOWN.Top = UserControl.Height / Screen.TwipsPerPixelY - btnDOWN.Height - 2
        SetScroller
        If dView = 0 Or dView = 3 Then picScr.Visible = True
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 2
    Else
        picScr.Visible = False
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    End If

Else
    CheckPath
    If qk = False Then RaiseEvent PathChange(ipath, Bpath)
    
    'create what user sees
    RaiseEvent DirSelect(ipath)
    
    
    
    If DoNotGenerate = True Then
        picOzadje.Cls
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
        picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelY - 2
        picScr.Visible = False
        picScrH.Visible = False
        Dim cc As String
        cc = "Error: This path is not valid!"
        
        picBuffText.Font.Bold = True
        picBuffText.Width = picBuffText.TextWidth(cc)
        picBuffText.Cls
        picBuffText.ForeColor = SelBorderColor
        picBuffText.Print cc
        
        picBuffText.Font.Bold = False
        
        BitBlt picOzadje.hDC, (picOzadje.Width - picBuffText.Width) / 2, (picOzadje.Height - picBuffText.Height) / 2, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
        picOzadje.Refresh
        ListCount = 0
        Exit Property
    End If
    
    picOzadje.Cls
    iTop = 0
    CheckPath
    If qPathExists = True Then
    File1.Path = dPath
    File1.Refresh
    
    If Right(File1.Path, 1) = "\" Then
        cPath = File1.Path
    Else
        cPath = File1.Path & "\"
    End If
    List1.Clear
    
    For ICNT = 0 To Dir1.ListCount - 1
        List1.AddItem Dir1.List(ICNT)
    
    Next ICNT

    For ICNT = 0 To File1.ListCount - 1
        If ExtensionIsValid(GetExtension(File1.List(ICNT))) = True Then
            List1.AddItem cPath & File1.List(ICNT)
        End If
    Next ICNT
    End If
    Dim K As Integer
    K = 0
    If Capitalize(DesktopFolder) = Capitalize(ipath) Then   'CHECK IF WE ARE ON THE DESKTOP
        List1.AddItem "#MYCOMPUTER", 0                      'IF WE ARE, WE HAVE TO CREATE ARRAY TO MYCOMPUTER AND MY DOCUMENTS
        List1.AddItem GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal"), 0    'MYDOCUMENTS FOLDER
        K = 2
    End If
    
    ListCount = List1.ListCount

    ReDim Preserve iFile(ListCount)
    
    For ICNT = 1 To ListCount
        iFile(ICNT).FullName = List1.List(ICNT - 1)
        iFile(ICNT).Selected = False
        iFile(ICNT).ShowDetails = False
        
        If List1.List(ICNT - 1) = "#MYCOMPUTER" Then
            iFile(ICNT).name = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "")
            If iFile(ICNT).name = "" Then iFile(ICNT).name = "My Computer"
        ElseIf List1.List(ICNT - 1) = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") Then
            iFile(ICNT).name = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}", "")
            If iFile(ICNT).name = "" Then iFile(ICNT).name = "My Documents"
        Else
            iFile(ICNT).name = Mid(iFile(ICNT).FullName, InStrRev(iFile(ICNT).FullName, "\") + 1)
        End If
        
        If ICNT <= Dir1.ListCount + K Then
            iFile(ICNT).Extension = "#FOLDER"
        Else
            iFile(ICNT).Extension = GetExtension(iFile(ICNT).name)
        End If
        
        
        If Icon.Count > 1 Then
        
            
            IconFound = False
            
            For IcNT2 = 1 To Icon.Count - 1
                If iFile(ICNT).Extension = Icon(IcNT2).Tag Then
                    IconFound = True
                    Exit For
                End If
            Next IcNT2
            
            If IconFound = True And iFile(ICNT).FullName <> "#MYCOMPUTER" And iFile(ICNT).FullName <> GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") And iFile(ICNT).Extension <> "LNK" And iFile(ICNT).Extension <> "EXE" And iFile(ICNT).Extension <> "ICO" Then  'for every file we find a sutible preloaded icon... except for the fileswith .lnk xtensions (srtcuts - they each have its own icon - according to the file they point) also .exe files have their own icons
                iFile(ICNT).IconIndex = IcNT2
            Else
                Load Icon(Icon.Count)
                Load IconS(IconS.Count)
                Load LIcon(LIcon.Count)
                Load LIconS(LIconS.Count)
                
                iFile(ICNT).IconIndex = Icon.Count - 1
                
                If iFile(ICNT).FullName = "#MYCOMPUTER" Then
                    Icon(Icon.Count - 1).Tag = "#MYCOMPUTER"
                Else
                    Icon(Icon.Count - 1).Tag = iFile(ICNT).Extension
                End If
                
                IconBuff.Height = 16
                IconBuff.Width = 16
                
                IconBuff.Cls
                IconBuff.BackColor = BackColor
                IconBuff.Picture = LoadPicture("")
            
            
                Dim hLargeIcon As Long
                Dim hSmallIcon As Long
            
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

                If iFile(ICNT).FullName = "#MYCOMPUTER" Then
                    ExtractIconEx "explorer.exe", 0, hLargeIcon, hSmallIcon, 1
                    DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
                ElseIf iFile(ICNT).FullName = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") Then
                    ExtractIconEx "mydocs.dll", 0, hLargeIcon, hSmallIcon, 1
                    DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
                Else
                    ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                End If
                

                
                IconBuff.Refresh
            
                Icon(Icon.Count - 1).Picture = IconBuff.Image
                
                IconBuff.Cls
                IconBuff.BackColor = SelBackColor
                IconBuff.Picture = LoadPicture("")
                
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

                If iFile(ICNT).FullName = "#MYCOMPUTER" Then
                    DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
                    DestroyIcon hSmallIcon
                ElseIf iFile(ICNT).FullName = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") Then
                    DrawIconEx IconBuff.hDC, 0, 0, hSmallIcon, 0, 0, 0, 0, DI_NORMAL
                    DestroyIcon hSmallIcon
                Else
                    ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                End If
                
                IconBuff.Refresh
                IconS(Icon.Count - 1).Picture = IconBuff.Image
                
                IconBuff.Height = 32
                IconBuff.Width = 32
                
                IconBuff.Cls
                IconBuff.BackColor = BackColor
                IconBuff.Picture = LoadPicture("")
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
            
                If iFile(ICNT).FullName = "#MYCOMPUTER" Then
                    DrawIconEx IconBuff.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
                ElseIf iFile(ICNT).FullName = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") Then
                    DrawIconEx IconBuff.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
                Else
                    ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                End If
                
                IconBuff.Refresh
            
                LIcon(Icon.Count - 1).Picture = IconBuff.Image
                
                IconBuff.Cls
                IconBuff.BackColor = SelBackColor
                IconBuff.Picture = LoadPicture("")
                iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
                If iFile(ICNT).FullName = "#MYCOMPUTER" Then
                    DrawIconEx IconBuff.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
                    DestroyIcon hLargeIcon
                ElseIf iFile(ICNT).FullName = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") Then
                    DrawIconEx IconBuff.hDC, 0, 0, hLargeIcon, 0, 0, 0, 0, DI_NORMAL
                    DestroyIcon hLargeIcon
                Else
                    ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
                End If
                
                IconBuff.Refresh
            
                LIconS(Icon.Count - 1).Picture = IconBuff.Image
            End If
        Else
            Load Icon(Icon.Count)
            Load IconS(IconS.Count)
            Load LIcon(LIcon.Count)
            Load LIconS(LIconS.Count)
            
            iFile(ICNT).IconIndex = Icon.Count - 1
            Icon(Icon.Count - 1).Tag = iFile(ICNT).Extension
            
            IconBuff.Height = 16
            IconBuff.Width = 16
            
            IconBuff.Cls
            IconBuff.BackColor = BackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            Icon(Icon.Count - 1).Picture = IconBuff.Image
            
            IconBuff.Cls
            IconBuff.BackColor = SelBackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            IconS(Icon.Count - 1).Picture = IconBuff.Image
            
            IconBuff.Height = 32
            IconBuff.Width = 32
            
            IconBuff.Cls
            IconBuff.BackColor = BackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            LIcon(Icon.Count - 1).Picture = IconBuff.Image
            
            IconBuff.Cls
            IconBuff.BackColor = SelBackColor
            IconBuff.Picture = LoadPicture("")
            iIcon = SHGetFileInfo(iFile(ICNT).FullName, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        
            ImageList_Draw iIcon, SHInfo.iIcon, IconBuff.hDC, 0, 0, ILD_TRANSPARENT
            IconBuff.Refresh
        
            LIconS(Icon.Count - 1).Picture = IconBuff.Image
            End If
    Next ICNT
    
   ' If Icon.count > ListCount + 1 Then
     '   For iCnt = ListCount + 2 To Icon.count - 1
    '        Unload Icon(iCnt)
    '        Unload IconS(iCnt)
    '    Next iCnt
    '
    'End If
    'Before we Genrate it we must check the size of the list (height - to set scrollers)
    DoNotGenerate = True
    
    If dView = 0 Then
        SelHeight = IconBuff.Height + 2 'NormalviewHeight
        selDetHeight = 64               '"Extendet" dView Mode
    End If
    
    If CheckHeight > UserControl.Height / Screen.TwipsPerPixelY Then
        btnDOWN.Top = UserControl.Height / Screen.TwipsPerPixelY - btnDOWN.Height - 2
        SetScroller
        If dView = 0 Or dView = 3 Then picScr.Visible = True
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 2
    Else
        picScr.Visible = False
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    End If

End If

DoNotGenerate = False
Generate

Dim isSelcted As Boolean
isSelcted = False
For ICNT = 1 To ListCount
    If iFile(ICNT).Selected = True And Len(iFile(ICNT).FullName) > 3 And iFile(ICNT).FullName <> "#MYCOMPUTER" Then
        isSelcted = True
    End If
Next ICNT
RaiseEvent DeletableItemSelected(isSelcted)

End Property

Public Sub Generate()

picOzadje.Cls
picOzadje.BackColor = BackColor
If DoNotGenerate = True Or dPath = "" Then GoTo CCEND

If dView = 0 Then
    SelHeight = 18 'NormalviewHeight
    selDetHeight = 64               '"Extendet" dView Mode
ElseIf dView = 1 Then
    SelHeight = 18 'NormalviewHeight
    selDetHeight = 0                'NoExtendetMode here
ElseIf dView = 2 Then
    SelHeight = 34 'NormalviewHeight
    selDetHeight = 0                'NoExtendetMode here
End If
    
Dim FromIndex As Long
Dim ToIndex As Long
Dim ICNT As Integer

If dView = 0 Then
    'IconSize = 16
    'dView = List
    'Allows ShowDetails
    
    If Not picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelX - 2 Then picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelX - 2
    picScrH.Visible = False
    
    IconBuff.Height = 16
    IconBuff.Width = 16
    
    FontSize = 10
    
    'SetHeight
    iHeight = CheckHeight
    
    DrawBackgrounds
    
    picBuffText.Height = picBuff.Height - 2
    
    ToIndex = ListCount
    Dim xTop As Integer
    Dim cHeight As Integer
    Dim cLeft As Integer
    Dim fTop As Long
    FromIndex = 0
    On Error GoTo loopend
    Do Until fTop > -iTop
        FromIndex = FromIndex + 1
        xTop = fTop + iTop
        If iFile(FromIndex).ShowDetails = True Then
    
            fTop = fTop + selDetHeight
    
        Else
            fTop = fTop + SelHeight
        End If
    Loop
loopend:

    If FromIndex < 1 Then FromIndex = 1
    fTop = 0
    For ICNT = 1 To ListCount
        If iFile(ICNT).ShowDetails = True Then
            fTop = fTop + selDetHeight
        Else
            fTop = fTop + SelHeight
        End If
        
        If fTop + iTop >= picOzadje.Height Then
            ToIndex = ICNT
            Exit For
        End If
    Next ICNT
    
    On Error Resume Next 'GoTo NoDraw

    For ICNT = FromIndex To ToIndex
    If ListCount > 0 Then
        If iFile(ICNT).ShowDetails = True Then
            picBuffText.Font.Size = 9
            picBuffText.Font.Bold = True
            
            If iFile(ICNT).Details = "" Then
                If iFile(ICNT).Extension = "#FOLDER" Then
                
                    Dim FOLDER
                    If Len(iFile(ICNT).FullName) < 4 Then
                        Set FOLDER = FSO.GetDrive(iFile(ICNT).FullName)
                        If FOLDER.IsReady = True Then
                            sqy = True
                            iFile(ICNT).Details = CAPTION_DRIVESIZE & GetSize(FOLDER.TotalSize)
                            iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_DRIVEFREESPACE & GetSize(FOLDER.FreeSpace)
                            iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_FILESYSTEM & FOLDER.FileSystem
                            iFile(ICNT).Details = iFile(ICNT).Details & "|"
                            sqy = False
                        Else
                            iFile(ICNT).Details = CAPTION_DEVICEUNAVAILABLE & "|||"
                        End If
                    Else
                        Set FOLDER = FSO.GetFolder(iFile(ICNT).FullName)
                        iFile(ICNT).Details = CAPTION_FOLDERPATH & FOLDER.Path
                        iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_FILESINFOLDER & FOLDER.Files.Count
                        iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_FOLDERSIZE & GetSize(FOLDER.Size)
                        iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_LASTACCESSED & Format(FOLDER.DateLastModified, "d/ mmmm yyyy") & ", " & Format(FOLDER.DateLastModified, "hh:nn:ss")
                    End If
                Else
                    iFile(ICNT).Details = CAPTION_FILETYPE & GetFileType(iFile(ICNT).FullName)
                    iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_FILESIZE & GetFileSize(iFile(ICNT).FullName)
                    iFile(ICNT).Details = iFile(ICNT).Details & "|" & CAPTION_LASTACCESSED & GetFileDate(iFile(ICNT).FullName)
                    iFile(ICNT).Details = iFile(ICNT).Details & "|"
                End If
            End If
            
            If iFile(ICNT).Selected = True Then
                IconBuff.Picture = LIconS(iFile(ICNT).IconIndex).Picture
                
                picBuff.Height = picExtendet.Height / 2
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picExtendet.hDC, 0, picExtendet.Height / 2, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                picBuffText.ForeColor = SelForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                picHideDetails.BackColor = picBuff.BackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = 57
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuffText.Font.Size = picBuffText.Font.Size - 2
                picBuffText.Font.Bold = True
                
                picBuffText.Width = picBuff.Width - cLeft - 8
                picBuffText.Height = picBuffText.TextHeight("gŽ") * 4
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                picBuffText.ForeColor = AlphaBlend(SelBackColor, SelForeColor, 110)
                
                Dim asd() As String
                asd() = Split(iFile(ICNT).Details, "|")
                
                picBuffText.Print asd(0)
                picBuffText.Print asd(1)
                picBuffText.Print asd(2)
                picBuffText.Print asd(3)
                
                picBuffText.Font.Size = picBuffText.Font.Size + 2
                picBuffText.Font.Bold = False
                cLeft = cLeft + 7
                BitBlt picBuff.hDC, cLeft, picBuffText.Height / 4 + 6, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY

                
                picBuffText.Height = picBuffText.TextHeight("gŽ")
                
                picBuff.Refresh
                
                
                
                'the following code "deletes" the upper and lower lines in multiple selections
                If ICNT > 1 Then
                    If iFile(ICNT - 1).Selected = True Then
                        picBuff.ForeColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                        picBuff.Line (56, 0)-(picBuff.Width - 1, 0)
                        If iFile(ICNT - 1).ShowDetails = True Then picBuff.ForeColor = SelBackColor
                        picBuff.Line (1, 0)-(56, 0)
                    End If
                End If
    
                If ICNT < ListCount Then
                    If iFile(ICNT + 1).Selected = True Then
                        picBuff.ForeColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                    
                    End If
                End If
                                
                BitBlt picOzadje.hDC, 0, xTop, picOzadje.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height
            Else
                IconBuff.Picture = LIcon(iFile(ICNT).IconIndex).Picture
                
                picBuff.Height = picExtendet.Height / 2
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picExtendet.hDC, 0, 0, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(BackColor, vbBlack, 245)
                picBuffText.ForeColor = ForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                picHideDetails.BackColor = BackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = 57
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuffText.Font.Size = picBuffText.Font.Size - 2
                picBuffText.Font.Bold = True
                
                picBuffText.Width = picBuff.Width - cLeft - 8
                picBuffText.Height = picBuffText.TextHeight("gŽ") * 4
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(BackColor, vbBlack, 245)
                picBuffText.ForeColor = AlphaBlend(BackColor, ForeColor, 110)
                
                asd() = Split(iFile(ICNT).Details, "|")
                
                picBuffText.Print asd(0)
                picBuffText.Print asd(1)
                picBuffText.Print asd(2)
                picBuffText.Print asd(3)
                
                picBuffText.Font.Size = picBuffText.Font.Size + 2
                picBuffText.Font.Bold = False
                cLeft = cLeft + 7
                BitBlt picBuff.hDC, cLeft, picBuffText.Height / 4 + 6, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY

                
                picBuffText.Height = picBuffText.TextHeight("gŽ")
                
                picBuff.Refresh

                'the following code "deletes" the upper and lower lines in multiple selections
                If ICNT > 1 Then
                    If iFile(ICNT - 1).Selected = False Then
                        If iFile(ICNT - 1).ShowDetails = True Then picBuff.ForeColor = AlphaBlend(BackColor, vbBlack, 245) Else picBuff.ForeColor = AlphaBlend(BackColor, vbBlack, 225)
                        picBuff.Line (0, 0)-(picBuff.Width, 0)
                        picBuff.ForeColor = BackColor
                        If iFile(ICNT - 1).ShowDetails = True Then picBuff.Line (0, 0)-(55, 0)
                    End If
                End If
    
                If ICNT < ListCount Then
                    If iFile(ICNT + 1).Selected = False Then
                        'picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                    
                    End If
                End If
                                
                BitBlt picOzadje.hDC, 0, xTop, picOzadje.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height

            End If
            
        picBuff.Height = SelHeight
        picBuffText.Font.Size = 8
        picBuffText.Font.Bold = False
        
        Else
        
            'to add an icon infront of the caption
            If iFile(ICNT).Selected = True Then
                IconBuff.Picture = IconS(iFile(ICNT).IconIndex).Picture
    
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, picBackground.Height / 2, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = SelBackColor
                picBuffText.ForeColor = SelForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                picDetails.BackColor = picBuff.BackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
                
                picBuff.ForeColor = SelBackColor
                
                'the following code "deletes" the upper and lower lines in multiple selections
                If ICNT > 1 Then
                    If iFile(ICNT - 1).Selected = True Then
                        picBuff.Line (1, 0)-(picBuff.Width - 1, 0)
                    
                    End If
                End If
    
                If ICNT < ListCount Then
                    If iFile(ICNT + 1).Selected = True Then
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                    
                    End If
                End If
                                
                BitBlt picOzadje.hDC, 0, xTop, picOzadje.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height
            Else
                IconBuff.Picture = Icon(iFile(ICNT).IconIndex).Picture
            
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, 0, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = BackColor
                picBuffText.ForeColor = ForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                
                picDetails.BackColor = picBuffText.BackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
    
    
                BitBlt picOzadje.hDC, 0, xTop, picOzadje.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height
            End If
        
        End If
    End If
    Next ICNT
NoDraw:
    
    picOzadje.Refresh
    
'_________________________________________________________________________'
'___________________________________000000________________________________'
'________________________________000___000________________________________'
'______________________________________000________________________________'
'______________________________________000________________________________'
'______________________________________000________________________________'
'______________________________________000________________________________'
'______________________________________000________________________________'
'________________________________000000000000000__________________________'
'_________________________________________________________________________'
ElseIf dView = 1 Then       'List with small icons
    IconBuff.Width = 16     'SmallIconsSize
    IconBuff.Height = 16
    
    If Not picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2 Then picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    
    If picScr.Visible = True Then picScr.Visible = False
    
    Dim startColumn As Integer
    Dim endColumn As Integer
    
    Dim xLeft As Long                   'Position left of the list
    
    xLeft = 0
    
    RowsInColumnCount = Int((UserControl.Height / Screen.TwipsPerPixelY - 2) / SelHeight) 'how manny lines can we draw in each column
    
    ColumnCount = ListCount / RowsInColumnCount             'calculate how many columns we have
    
    If ColumnCount * RowsInColumnCount < ListCount Then
        ColumnCount = ColumnCount + 1
    End If
    
    startColumn = 1
    endColumn = ColumnCount
    
    'to get text width
    Dim TextWidth As Integer
    
    For ICNT = 1 To ListCount
        If TextWidth < picBuffText.TextWidth(iFile(ICNT).name) Then
            TextWidth = picBuffText.TextWidth(iFile(ICNT).name)
        End If
    Next ICNT
    
    'SetOneColumn Width
    ColumnWidth = TextWidth + 4 + IconBuff.Width 'textwidth + borders and spaces and icon width
      
    
    Dim XwIDTH As Integer
    XwIDTH = 0

    XwIDTH = ColumnWidth * ColumnCount
    
    If XwIDTH > picOzadje.Width Then
        picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelX - 2 - picScrH.Height
        picScrH.Top = UserControl.Height / Screen.TwipsPerPixelY - picScrH.Height - 1
        picScrH.Width = picOzadje.Width
        picScrH.Visible = True
        
        If iLeft + ColumnWidth * ColumnCount < picOzadje.Width Then iLeft = picOzadje.Width - ColumnWidth * ColumnCount
        
        'we have to do this again since the height has changed
        RowsInColumnCount = Int(picOzadje.Height / SelHeight)   'how manny lines can we draw in each column
        
        'ColumnCount = ListCount / RowsInColumnCount             'calculate how many columns we have
        
        'If ColumnCount * RowsInColumnCount < ListCount Then
       '     ColumnCount = ColumnCount + 1
        'End If
    
        For ICNT = 1 To ColumnCount
            If iLeft + ICNT * ColumnWidth > 0 Then
                Exit For
            End If
        Next ICNT
        
        If ICNT > 1 Then startColumn = ICNT - 1 Else startColumn = 1 ' We pass this so we do not need to generate invisible items
        
        For ICNT = 1 To ColumnCount
            If iLeft + ICNT * ColumnWidth > picOzadje.Width Then
                Exit For
            End If
        Next ICNT
        If ICNT > 1 Then endColumn = ICNT
        
    Else
        iLeft = 0
        picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelX - 2                   'and when it's hidden
        picScrH.Visible = False
    End If
      
    picBuff.Width = ColumnWidth
    If picBackground.Width <> picBuff.Width Or picBackground.Height <> picBuff.Height Then
        DrawBackgrounds
    End If
    
    picBuff.Height = SelHeight
    picBuffText.Height = picBuff.Height - 2
    
    Dim StartFile As Integer
    Dim EndFile As Integer
    
    Dim iCntC As Integer 'to go trough columns
    For ICNT = 1 To ColumnCount
        If (ICNT) * ColumnWidth + iLeft >= 0 Then
            startColumn = ICNT
            Exit For
        End If
    Next ICNT
    
    For ICNT = 1 To ColumnCount
        If (ICNT) * ColumnWidth + iLeft >= picOzadje.Width Then
            endColumn = ICNT
            Exit For
        End If
    Next ICNT
    
    If endColumn < 1 Or endColumn > ColumnCount Then endColumn = ColumnCount

    For iCntC = startColumn To endColumn
    
        xLeft = iLeft + (iCntC - 1) * ColumnWidth
        
        StartFile = (iCntC - 1) * RowsInColumnCount + 1
        EndFile = (iCntC) * RowsInColumnCount
        If EndFile > ListCount Then EndFile = ListCount
        For ICNT = StartFile To EndFile 'to go troug each line in a column
        
            If iFile(ICNT).Selected = True Then
                IconBuff.Picture = IconS(iFile(ICNT).IconIndex).Picture
                
                picBuff.Cls
                
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, picBackground.Height / 2, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                
                picBuffText.BackColor = SelBackColor
                picBuffText.ForeColor = SelForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                
                picDetails.BackColor = picBuffText.BackColor
                                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
    
                'Now we deleete borders of selections where selections touch eachothers (looks cool;p)
                xTop = (ICNT - (iCntC - 1) * RowsInColumnCount - 1) * SelHeight
                
                picBuff.ForeColor = SelBackColor
                
                'TOP AND BOTTOM OF SELECTION
                Dim qBOTTOM As Boolean
                Dim qTOP As Boolean
                qBOTTOM = False
                qTOP = False
                
                If ICNT > 1 Then
                    If iFile(ICNT - 1).Selected = True And xTop > 0 Then
                        picBuff.Line (1, 0)-(picBuff.Width - 1, 0)
                        qTOP = True
                    End If
                
                End If

                If ICNT < ListCount Then
                    If iFile(ICNT + 1).Selected = True And xTop < (RowsInColumnCount - 1) * SelHeight Then
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                        qBOTTOM = True
                    End If
                
                End If
                
                'LEFT AND RIGHT
                Dim Y1 As Integer
                Dim Y2 As Integer
                
                If qTOP = True Then Y1 = 0 Else Y1 = 1
                If qBOTTOM = True Then Y2 = picBuff.Height Else Y2 = picBuff.Height - 1
                
                If ICNT + RowsInColumnCount <= ListCount Then
                    If iFile(ICNT + RowsInColumnCount).Selected = True Then
                        If qTOP = True And iFile(ICNT + RowsInColumnCount - 1).Selected = False Then ' FOR CORNERS
                            Y1 = 1
                        End If
                        
                        If (ICNT + RowsInColumnCount + 1) <= ListCount Then
                            If qBOTTOM = True And iFile(ICNT + RowsInColumnCount + 1).Selected = False Then ' FOR CORNERS
                                Y2 = picBuff.Height - 1
                            End If
                        Else
                            Y2 = picBuff.Height - 1 ' hard to explain - try to select all and delete this line and you might see the problem:p
                        End If
                        
                        picBuff.Line (picBuff.Width - 1, Y1)-(picBuff.Width - 1, Y2)
                    End If
                End If
                
                If qTOP = True Then Y1 = 0 Else Y1 = 1
                If qBOTTOM = True Then Y2 = picBuff.Height Else Y2 = picBuff.Height - 1

                
                If ICNT - RowsInColumnCount > 0 Then
                    If iFile(ICNT - RowsInColumnCount).Selected = True Then
                        If qTOP = True And iFile(ICNT - RowsInColumnCount - 1).Selected = False Then ' FOR CORNERS
                            Y1 = 1
                        End If
                        
                        If (ICNT - RowsInColumnCount + 1) <= ListCount Then
                            If qBOTTOM = True And iFile(ICNT - RowsInColumnCount + 1).Selected = False Then ' FOR CORNERS
                                Y2 = picBuff.Height - 1
                            End If
                        End If
                        picBuff.Line (0, Y1)-(0, Y2)
                    End If
                End If
                
                'NOW DRAW THE ITEM TO THE VIEW PICTURE BOX
                BitBlt picOzadje.hDC, xLeft, xTop, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY

            Else
                IconBuff.Picture = Icon(iFile(ICNT).IconIndex).Picture
                
                picBuff.Cls
                
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, 0, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                
                picBuffText.BackColor = BackColor
                picBuffText.ForeColor = ForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                
                picDetails.BackColor = picBuffText.BackColor
                                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
    
                xTop = (ICNT - (iCntC - 1) * RowsInColumnCount - 1) * SelHeight
                BitBlt picOzadje.hDC, xLeft, xTop, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY

            End If
        Next ICNT
    Next iCntC
    
    If picScrH.Visible = True Then
        btnRIGHT.Left = UserControl.Width / Screen.TwipsPerPixelX - 2 - btnRIGHT.Width
        SetScrollerH
    End If
    
'_________________________________________________________________________'
'___________________________________000000________________________________'
'________________________________000______000_____________________________'
'_________________________________________000_____________________________'
'_________________________________________000_____________________________'
'______________________________________000________________________________'
'___________________________________000___________________________________'
'________________________________000______________________________________'
'________________________________000000000000__________________________'
'_________________________________________________________________________'
ElseIf dView = 2 Then       'List with large icons
    IconBuff.Width = 32     'SmallIconsSize
    IconBuff.Height = 32
    
    If Not picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2 Then picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    
    If picScr.Visible = True Then picScr.Visible = False
    
    'Dim StartColumn As Integer
    'Dim EndColumn As Integer
    
    'Dim RowsInColumnCount As Integer    'how many lines can go into one column
    'Dim xLeft As Long                   'Position left of the list
    
    xLeft = 0
    
    RowsInColumnCount = Int((UserControl.Height / Screen.TwipsPerPixelX - 2) / SelHeight) 'how manny lines can we draw in each column
    
    ColumnCount = ListCount / RowsInColumnCount             'calculate how many columns we have
    
    If ColumnCount * RowsInColumnCount < ListCount Then
        ColumnCount = ColumnCount + 1
    End If

    startColumn = 1
    endColumn = ColumnCount
    
    'to get text width
    'Dim TextWidth As Integer
    
    For ICNT = 1 To ListCount
        If TextWidth < picBuffText.TextWidth(iFile(ICNT).name) Then
            TextWidth = picBuffText.TextWidth(iFile(ICNT).name)
        End If
    Next ICNT
    
    'SetOneColumn Width
    'Dim ColumnWidth As Integer
    ColumnWidth = TextWidth + 4 + IconBuff.Width 'textwidth + borders and spaces and icon width
      
    'Dim XwIDTH As Integer
    XwIDTH = 0

    XwIDTH = ColumnWidth * ColumnCount
    
    If XwIDTH > picOzadje.Width Then
        picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelX - 2 - picScrH.Height
        picScrH.Top = UserControl.Height / Screen.TwipsPerPixelY - picScrH.Height - 1
        picScrH.Width = picOzadje.Width
        picScrH.Visible = True
        
        If iLeft + ColumnWidth * ColumnCount < picOzadje.Width Then iLeft = picOzadje.Width - ColumnWidth * ColumnCount

        'we have to do this again since the height has changed
        RowsInColumnCount = Int(picOzadje.Height / SelHeight)   'how manny lines can we draw in each column
        
        ColumnCount = ListCount / RowsInColumnCount             'calculate how many columns we have
        
        If ColumnCount * RowsInColumnCount < ListCount Then
            ColumnCount = ColumnCount + 1
        End If
    
        For ICNT = 1 To ColumnCount
            If iLeft + ICNT * ColumnWidth > 0 Then
                Exit For
            End If
        Next ICNT
        
        If ICNT > 1 Then startColumn = ICNT - 1 Else startColumn = 1 ' We pass this so we do not need to generate invisible items
        
        For ICNT = 1 To ColumnCount
            If iLeft + ICNT * ColumnWidth > picOzadje.Width Then
                Exit For
            End If
        Next ICNT
        If ICNT > 1 Then endColumn = ICNT
        
    Else
        iLeft = 0
        picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelX - 2                   'and when it's hidden
        picScrH.Visible = False
    End If
      
    picBuff.Width = ColumnWidth
    If picBackground.Width <> picBuff.Width Or picBackground.Height <> picBuff.Height Then
        DrawBackgrounds
    End If
    
    picBuff.Height = SelHeight
    picBuffText.Height = picBuffText.TextHeight("Žg")
    
    'Dim StartFile As Integer
    'Dim EndFile As Integer
    
    'Dim iCntC As Integer 'to go trough columns
   ' Dim startColumn As Integer
    'Dim endColumn As Integer
    
    For ICNT = 1 To ColumnCount
        If (ICNT) * ColumnWidth + iLeft >= 0 Then
            startColumn = ICNT
            Exit For
        End If
    Next ICNT
    
    For ICNT = 1 To ColumnCount
        If (ICNT) * ColumnWidth + iLeft >= picOzadje.Width Then
            endColumn = ICNT
            Exit For
        End If
    Next ICNT
    
    If endColumn < 1 Or endColumn > ColumnCount Then endColumn = ColumnCount

    For iCntC = startColumn To endColumn
    
        xLeft = iLeft + (iCntC - 1) * ColumnWidth
        
        StartFile = (iCntC - 1) * RowsInColumnCount + 1
        EndFile = (iCntC) * RowsInColumnCount
        If EndFile > ListCount Then EndFile = ListCount
        For ICNT = StartFile To EndFile 'to go troug each line in a column
        
            If iFile(ICNT).Selected = True Then
                IconBuff.Picture = LIconS(iFile(ICNT).IconIndex).Picture
                
                picBuff.Cls
                
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, picBackground.Height / 2, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                
                picBuffText.BackColor = SelBackColor
                picBuffText.ForeColor = SelForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                
                picDetails.BackColor = picBuffText.BackColor
                                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picBuffText.Height) / 2, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
    
                'Now we deleete borders of selections where selections touch eachothers (looks cool;p)
                xTop = (ICNT - (iCntC - 1) * RowsInColumnCount - 1) * SelHeight
                
                picBuff.ForeColor = SelBackColor
                
                'TOP AND BOTTOM OF SELECTION
                'Dim qBOTTOM As Boolean
                'Dim qTOP As Boolean
                qBOTTOM = False
                qTOP = False
                
                If ICNT > 1 Then
                    If iFile(ICNT - 1).Selected = True And xTop > 0 Then
                        picBuff.Line (1, 0)-(picBuff.Width - 1, 0)
                        qTOP = True
                    End If
                
                End If

                If ICNT < ListCount Then
                    If iFile(ICNT + 1).Selected = True And xTop < (RowsInColumnCount - 1) * SelHeight Then
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                        qBOTTOM = True
                    End If
                
                End If
                
                'LEFT AND RIGHT
                'Dim Y1 As Integer
                'Dim Y2 As Integer
                
                If qTOP = True Then Y1 = 0 Else Y1 = 1
                If qBOTTOM = True Then Y2 = picBuff.Height Else Y2 = picBuff.Height - 1
                
                If ICNT + RowsInColumnCount <= ListCount Then
                    If iFile(ICNT + RowsInColumnCount).Selected = True Then
                        If qTOP = True And iFile(ICNT + RowsInColumnCount - 1).Selected = False Then ' FOR CORNERS
                            Y1 = 1
                        End If
                        
                        If (ICNT + RowsInColumnCount + 1) <= ListCount Then
                            If qBOTTOM = True And iFile(ICNT + RowsInColumnCount + 1).Selected = False Then ' FOR CORNERS
                                Y2 = picBuff.Height - 1
                            End If
                        Else
                            Y2 = picBuff.Height - 1 ' hard to explain - try to select all and delete this line and you might see the problem:p
                        End If
                        
                        picBuff.Line (picBuff.Width - 1, Y1)-(picBuff.Width - 1, Y2)
                    End If
                End If
                
                If qTOP = True Then Y1 = 0 Else Y1 = 1
                If qBOTTOM = True Then Y2 = picBuff.Height Else Y2 = picBuff.Height - 1

                
                If ICNT - RowsInColumnCount > 0 Then
                    If iFile(ICNT - RowsInColumnCount).Selected = True Then
                        If qTOP = True And iFile(ICNT - RowsInColumnCount - 1).Selected = False Then ' FOR CORNERS
                            Y1 = 1
                        End If
                        
                        If (ICNT - RowsInColumnCount + 1) <= ListCount Then
                            If qBOTTOM = True And iFile(ICNT - RowsInColumnCount + 1).Selected = False Then ' FOR CORNERS
                                Y2 = picBuff.Height - 1
                            End If
                        End If
                        picBuff.Line (0, Y1)-(0, Y2)
                    End If
                End If
                
                'NOW DRAW THE ITEM TO THE VIEW PICTURE BOX
                BitBlt picOzadje.hDC, xLeft, xTop, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY

            Else
                IconBuff.Picture = LIcon(iFile(ICNT).IconIndex).Picture
                picBuff.Cls
                
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, 0, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                picBuffText.Cls
                
                picBuffText.BackColor = BackColor
                picBuffText.ForeColor = ForeColor
                picBuffText.Print iFile(ICNT).name
                
                cLeft = 1
                
                picDetails.BackColor = picBuffText.BackColor
                                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picBuffText.Height) / 2, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
    
                xTop = (ICNT - (iCntC - 1) * RowsInColumnCount - 1) * SelHeight
                BitBlt picOzadje.hDC, xLeft, xTop, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
            End If
        Next ICNT
    Next iCntC
    
    If picScrH.Visible = True Then
        btnRIGHT.Left = UserControl.Width / Screen.TwipsPerPixelX - 2 - btnRIGHT.Width
        SetScrollerH
    End If
    
'_________________________________________________________________________'
'___________________________________000000________________________________'
'________________________________000______000_____________________________'
'_________________________________________000_____________________________'
'_________________________________________000_____________________________'
'______________________________________000________________________________'
'_________________________________________000_____________________________'
'_________________________________________000_____________________________'
'________________________________000______000_____________________________'
'___________________________________000000________________________________'
'_________________________________________________________________________'
ElseIf dView = 3 Then       'List with large icons
    IconBuff.Width = 32     'SmallIconsSize
    IconBuff.Height = 32
    
    picBuff.Width = IconBuff.Width * 3
    picBuffText.Height = picBuffText.TextHeight("Žg")
    picBuff.Height = IconBuff.Height + 5 + picBuffText.Height * 2
    
    If Not picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelY - 2 Then picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelY - 2
    
    If picScrH.Visible = True Then picScrH.Visible = False
    
    xTop = 0
    
    RowsInColumnCount = Int((UserControl.Width / Screen.TwipsPerPixelX - 2) / (picBuff.Width))  'how manny lines can we draw in each column
    
    ColumnCount = ListCount / RowsInColumnCount             'calculate how many columns we have
    
    If ColumnCount * RowsInColumnCount < ListCount Then
        ColumnCount = ColumnCount + 1
    End If

    startColumn = 1
    endColumn = ColumnCount
    
    'to get text width
    'Dim TextWidth As Integer
    
    For ICNT = 1 To ListCount
        'If TextWidth < picBuffText.TextWidth(iFile(iCnt).name) Then
        '    TextWidth = picBuffText.TextWidth(iFile(iCnt).name)
        'End If
    Next ICNT
    
    'SetOneColumn Width
    'Dim ColumnWidth As Integer
    ColumnWidth = picBuff.Height
      
    'Dim XwIDTH As Integer
    XwIDTH = 0

    XwIDTH = ColumnWidth * ColumnCount
    
    If XwIDTH > picOzadje.Height Then
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2 - picScr.Width
        picScr.Left = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 1
        picScr.Height = picOzadje.Height
        picScr.Visible = True
        
        'If iLeft + ColumnWidth * ColumnCount < picOzadje.Width Then iLeft = picOzadje.Width - ColumnWidth * ColumnCount

        'we have to do this again since the height has changed
        RowsInColumnCount = Int(picOzadje.Width / picBuff.Width)    'how manny lines can we draw in each column
        
        ColumnCount = ListCount / RowsInColumnCount             'calculate how many columns we have
        
        If ColumnCount * RowsInColumnCount < ListCount Then
            ColumnCount = ColumnCount + 1
        End If
    
        For ICNT = 1 To ColumnCount
            If iTop + ICNT * ColumnWidth > 0 Then
                Exit For
            End If
        Next ICNT
        
        If ICNT > 1 Then startColumn = ICNT - 1 Else startColumn = 1 ' We pass this so we do not need to generate invisible items
        
        For ICNT = 1 To ColumnCount
            If iTop + ICNT * ColumnWidth > picOzadje.Height Then
                Exit For
            End If
        Next ICNT
        If ICNT > 1 Then endColumn = ICNT
        
    Else
        iLeft = 0
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelY - 2                   'and when it's hidden
        picScr.Visible = False
    End If
    
    'picBuff.Width = ColumnWidth
    If picBackground.Width <> picBuff.Width Or picBackground.Height <> picBuff.Height Then
        DrawBackgrounds
    End If
    'picBuff.Height = SelHeight
    
    
    'Dim StartFile As Integer
    'Dim EndFile As Integer
    
    'Dim iCntC As Integer 'to go trough columns
   ' Dim startColumn As Integer
    'Dim endColumn As Integer
    
    For ICNT = 1 To ColumnCount
        If (ICNT) * ColumnWidth + iTop >= 0 Then
            startColumn = ICNT
            Exit For
        End If
    Next ICNT
    
    For ICNT = 1 To ColumnCount
        If (ICNT) * ColumnWidth + iTop >= picOzadje.Height Then
            endColumn = ICNT
            Exit For
        End If
    Next ICNT
    
    If endColumn < 1 Or endColumn > ColumnCount Then endColumn = ColumnCount

    For iCntC = startColumn To endColumn
    
        xTop = iTop + (iCntC - 1) * picBuff.Height
        
        StartFile = (iCntC - 1) * RowsInColumnCount + 1
        EndFile = (iCntC) * RowsInColumnCount
        If EndFile > ListCount Then EndFile = ListCount
        For ICNT = StartFile To EndFile 'to go troug each line in a column
            
            Dim T1 As String
            Dim T2 As String
            Dim iPos As Integer
            Dim K As Integer
            Dim T2Len As Integer
            
            If iFile(ICNT).Selected = True Then
                IconBuff.Picture = LIconS(iFile(ICNT).IconIndex).Picture
                picBuff.Cls
                
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, picBackground.Height / 2, SRCCOPY
                
                
                picBuffText.Cls
                picBuffText.BackColor = SelBackColor
                picBuffText.ForeColor = SelForeColor
                

                    
                If picBuffText.TextWidth(iFile(ICNT).name) > picBuff.Width - 2 Then
                    T1 = ""
                    T2 = ""
                    
                    iPos = 0
                    T1 = iFile(ICNT).name
                    For K = 1 To Len(T1)
                        iPos = InStrRev(T1, " ")
                        
                        If iPos > 0 Then T1 = Left(T1, iPos - 1)
                        
                        If picBuffText.TextWidth(T1) <= picBuff.Width - 2 Then
                            Exit For
                        End If
                        
                    Next K
                    
                    If iPos = 0 Then
                        T1 = iFile(ICNT).name
                        For K = 1 To Len(iFile(ICNT).name)
                            T1 = Left(T1, Len(T1) - 1)
                            If picBuffText.TextWidth(T1 & "...") <= picBuff.Width - 2 Then
                                Exit For
                            End If
                        Next K
                        T1 = T1 & "..."
                    Else
                        T2 = Mid(iFile(ICNT).name, Len(T1) + 2)
                        
                        If picBuffText.TextWidth(T2) > picBuff.Width - 2 Then
                            T2Len = Len(T2)
                            
                            For K = 1 To T2Len
                                T2 = Left(T2, Len(T2) - 1)
                                If picBuffText.TextWidth(T2 & "...") <= picBuff.Width - 2 Then
                                    Exit For
                                End If
                            Next K
                            
                            T2 = T2 & "..."
                        End If
                    End If
    
                        
                    picBuffText.Width = picBuffText.TextWidth(T1)
                    picBuffText.Print T1
                    BitBlt picBuff.hDC, Int((picBuff.Width - picBuffText.Width) / 2), IconBuff.Height + 3, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                    
                    picBuffText.Cls
                    picBuffText.Width = picBuffText.TextWidth(T2)
                    picBuffText.Print T2
                    BitBlt picBuff.hDC, (picBuff.Width - picBuffText.Width) / 2, IconBuff.Height + 3 + picBuffText.Height, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY

                Else
                    picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                    picBuffText.Print iFile(ICNT).name
                    BitBlt picBuff.hDC, (picBuff.Width - picBuffText.Width) / 2, IconBuff.Height + 3, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                End If
            
                BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY

                xLeft = (ICNT - (iCntC - 1) * RowsInColumnCount - 1) * picBuff.Width
                                
                Dim qLeft As Boolean
                Dim qRight As Boolean
                Dim X1 As Integer
                Dim X2 As Integer
                
                qLeft = False
                qRight = False
                
                picBuff.ForeColor = SelBackColor
                
                'LEFT AND RIGHT
                If ICNT > 1 Then
                    If iFile(ICNT - 1).Selected = True And xLeft > 0 Then
                        picBuff.Line (0, 1)-(0, picBuff.Height - 1)
                        qLeft = True
                    End If
                End If
                
                If ICNT < ListCount Then
                    If iFile(ICNT + 1).Selected = True And xLeft < (RowsInColumnCount - 1) * picBuff.Width Then
                        picBuff.Line (picBuff.Width - 1, 1)-(picBuff.Width - 1, picBuff.Height - 1)
                        qRight = True
                    End If
                End If
                
                If qLeft = True Then X1 = 0 Else X1 = 1
                If qRight = True Then X2 = picBuff.Width Else X2 = picBuff.Width - 1

                If ICNT + RowsInColumnCount <= ListCount Then
                    If iFile(ICNT + RowsInColumnCount).Selected = True Then
                        If qLeft = True And iFile(ICNT + RowsInColumnCount - 1).Selected = False Or xLeft = 0 Then ' FOR CORNERS
                            X1 = 1
                        End If
                        
                        If (ICNT + RowsInColumnCount + 1) <= ListCount Then
                            If qRight = True And iFile(ICNT + RowsInColumnCount + 1).Selected = False Or ICNT = iCntC * RowsInColumnCount Then ' FOR CORNERS
                                X2 = picBuff.Width - 1
                            End If
                        Else
                            X2 = picBuff.Width - 1 ' hard to explain - try to select all and delete this line and you might see the problem:p
                        End If
                        
                        picBuff.Line (X1, picBuff.Height - 1)-(X2, picBuff.Height - 1)
                    End If
                End If
                
                If qLeft = True Then X1 = 0 Else X1 = 1
                If qRight = True Then X2 = picBuff.Width Else X2 = picBuff.Width - 1

                If ICNT - RowsInColumnCount > 0 Then
                    If iFile(ICNT - RowsInColumnCount).Selected = True Then
                        If qLeft = True And iFile(ICNT - RowsInColumnCount - 1).Selected = False Or xLeft = 0 Then  ' FOR CORNERS
                            X1 = 1
                        End If
                        
                        If (ICNT - RowsInColumnCount + 1) <= ListCount Then
                            If qRight = True And iFile(ICNT - RowsInColumnCount + 1).Selected = False Or ICNT = iCntC * RowsInColumnCount Then  ' FOR CORNERS
                                X2 = picBuff.Width - 1
                            End If
                        End If
                        picBuff.Line (X1, 0)-(X2, 0)
                    End If
                End If
                
                BitBlt picOzadje.hDC, xLeft, xTop, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
            Else
                IconBuff.Picture = LIcon(iFile(ICNT).IconIndex).Picture
                picBuff.Cls
                
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, 0, SRCCOPY
                
                
                picBuffText.Cls
                picBuffText.BackColor = BackColor
                picBuffText.ForeColor = ForeColor
                
                If picBuffText.TextWidth(iFile(ICNT).name) > picBuff.Width - 2 Then
                    T1 = ""
                    T2 = ""
                    
                    iPos = 0
                    T1 = iFile(ICNT).name
                    For K = 1 To Len(T1)
                        iPos = InStrRev(T1, " ")
                        
                        If iPos > 0 Then T1 = Left(T1, iPos - 1)
                        
                        If picBuffText.TextWidth(T1) <= picBuff.Width - 2 Then
                            Exit For
                        End If
                        
                    Next K
                    
                    If iPos = 0 Then
                        T1 = iFile(ICNT).name
                        For K = 1 To Len(iFile(ICNT).name)
                            T1 = Left(T1, Len(T1) - 1)
                            If picBuffText.TextWidth(T1 & "...") <= picBuff.Width - 2 Then
                                Exit For
                            End If
                        Next K
                        T1 = T1 & "..."
                    Else
                        T2 = Mid(iFile(ICNT).name, Len(T1) + 2)
                        
                        If picBuffText.TextWidth(T2) > picBuff.Width - 2 Then
                            T2Len = Len(T2)
                            For K = 1 To T2Len
                                T2 = Left(T2, Len(T2) - 1)
                                If picBuffText.TextWidth(T2 & "...") <= picBuff.Width - 2 Then
                                    Exit For
                                End If
                            Next K
                            T2 = T2 & "..."
                        End If
                    End If
    
                        
                    picBuffText.Width = picBuffText.TextWidth(T1)
                    picBuffText.Print T1
                    BitBlt picBuff.hDC, (picBuff.Width - picBuffText.Width) / 2, IconBuff.Height + 3, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                    
                    picBuffText.Cls
                    picBuffText.Width = picBuffText.TextWidth(T2)
                    picBuffText.Print T2
                    BitBlt picBuff.hDC, (picBuff.Width - picBuffText.Width) / 2, IconBuff.Height + 3 + picBuffText.Height, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY

                Else
                    picBuffText.Width = picBuffText.TextWidth(iFile(ICNT).name)
                    picBuffText.Print iFile(ICNT).name
                    BitBlt picBuff.hDC, (picBuff.Width - picBuffText.Width) / 2, IconBuff.Height + 3, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                End If
            
                BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY

                xLeft = (ICNT - (iCntC - 1) * RowsInColumnCount - 1) * picBuff.Width
                BitBlt picOzadje.hDC, xLeft, xTop, picBuff.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
            End If
        Next ICNT
    Next iCntC
    
    SetScroller
End If

CCEND:
picOzadje.Refresh

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

Public Function GetExtension(FileName As String) As String
'This gets file's extension... It also capitalizes it so mp3 and MP3 are the sam etc...
GetExtension = Capitalize(Mid(FileName, InStrRev(FileName, ".") + 1))

End Function

Private Sub CheckPath()
On Error GoTo handle
Dir1.Path = dPath ' JUST AN EASY WAY TO SEE IF THE PATH EXISTS
                 ' IF IT DOES NOT, DIR1 WILL RETURN AN ERROR
qPathExists = True
'DoNotGenerate = False
Exit Sub
handle:
RaiseEvent ErrorOcured(vbError) 'vberror = 10 if the path is not found etc.
'DoNotGenerate = True
qPathExists = False

End Sub

Private Function ExtensionIsValid(Extension As String) As Boolean
ExtensionIsValid = False
Dim ICNT As Integer

For ICNT = 0 To lstExtensions.ListCount - 1
    If lstExtensions.List(ICNT) = Extension Or lstExtensions.List(ICNT) = "*.*" Then
        ExtensionIsValid = True
    End If

Next ICNT

End Function

Public Property Let filter(filter As String)
lstExtensions.Clear

Dim ICNT As Integer
Dim asd() As String
asd() = Split(filter, "; ")

Dim K As String

For ICNT = 0 To UBound(asd)
    If asd(ICNT) = "*.*" Then
        lstExtensions.AddItem "*.*"
    Else
        lstExtensions.AddItem Capitalize(Mid$(asd(ICNT), 3))
    End If
Next ICNT
If DoNotGenerate = False Then Refresh

End Property

Public Property Get filter() As String
filter = UserControl.lstExtensions.List(0)

End Property

Public Sub Resize()
SetScroller
SetScrollerH
Generate

End Sub

Private Sub SetUpMnuExtendet(index As Long)
Dim ICNT As Integer
Dim cc As cMemDC

'SETS BUTTONS
For ICNT = 1 To btnMenuExtendet.Count - 1
    Unload btnMenuExtendet(ICNT)
Next ICNT

Dim tWidth As Integer
picMnuText.Font.Bold = False
tWidth = picMnuText.TextWidth(EXTMNU_EXPANDALL)
If tWidth < picMnuText.TextWidth(EXTMNU_UNEXPANDALL) Then tWidth = picMnuText.TextWidth(EXTMNU_UNEXPANDALL)

picMnuText.Font.Bold = True
If iFile(index).ShowDetails = False Then
    If tWidth < picMnuText.TextWidth(EXTMNU_EXPAND) Then tWidth = picMnuText.TextWidth(EXTMNU_EXPAND)
Else
    If tWidth < picMnuText.TextWidth(EXTMNU_UNEXPAND) Then tWidth = picMnuText.TextWidth(EXTMNU_UNEXPAND)
End If

tWidth = tWidth + 5
For ICNT = 1 To 3
    picMnuText.Font.Bold = False
    Load btnMenuExtendet(ICNT)
    btnMenuExtendet(ICNT).Tag = index
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
    SecondColor = AlphaBlend(vb3DFace, vb3DDKShadow, 240)

    Dim IcNT2 As Integer
    picBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 40)
    For IcNT2 = 1 To picBuff3.Width
        picBuff3.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * IcNT2 / picBuff3.Width))
        picBuff3.Line (IcNT2 - 1, 0)-(IcNT2 - 1, picBuff3.Height)
        picBuff3.Refresh
    Next IcNT2

    BitBlt picBuff.hDC, 0, 0, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    
    picBuff3.Cls
    picBuff3.Picture = MnuIconExtendet(ICNT - 1)

    picBuff3.Cls
    picBuff3.Picture = LoadPicture("")
    picMnuText.Width = tWidth
    picBuff3.BackColor = picBuff.BackColor
    
    picMnuText.Height = picMnuText.TextHeight("Žg")
    picMnuText.BackColor = picBuff.BackColor
    picMnuText.ForeColor = &H80000007
    picMnuText.Cls
    
    If ICNT = 1 Then
        picMnuText.Font.Bold = True
        If iFile(index).ShowDetails = False Then
            picMnuText.Print EXTMNU_EXPAND
        Else
            picMnuText.Print EXTMNU_UNEXPAND
        End If
    ElseIf ICNT = 2 Then
        picMnuText.Print EXTMNU_EXPANDALL
    ElseIf ICNT = 3 Then
        picMnuText.Print EXTMNU_UNEXPANDALL
    End If
    
    BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picMnuText.Height) / 2, picMnuText.Width, picMnuText.Height, picMnuText.hDC, 0, 0, SRCCOPY
    
    Set btnMenuExtendet(ICNT).NormalImage = picBuff.Image
    
    picBuff.Cls
    picBuff.ForeColor = vbHighlight
    picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.Picture = MnuIconExtendet(ICNT - 1)
    
    BitBlt picBuff.hDC, (picBuff.Height - 18) / 2, (picBuff.Height - 18) / 2, 18, 18, picBuff3.hDC, 0, 0, SRCCOPY

    picMnuText.Cls
    picBuff3.Picture = LoadPicture("")
    picMnuText.Width = tWidth
    picMnuText.BackColor = picBuff.BackColor
    picMnuText.ForeColor = &H80000007
    
    If ICNT = 1 Then
        If iFile(index).ShowDetails = False Then
            picMnuText.Print EXTMNU_EXPAND
        Else
            picMnuText.Print EXTMNU_UNEXPAND
        End If
    ElseIf ICNT = 2 Then
        picMnuText.Print EXTMNU_EXPANDALL
    ElseIf ICNT = 3 Then
        picMnuText.Print EXTMNU_UNEXPANDALL
    End If
    
    BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picMnuText.Height) / 2, picMnuText.Width, picMnuText.Height, picMnuText.hDC, 0, 0, SRCCOPY

    Set btnMenuExtendet(ICNT).FocusedImage = picBuff.Image
    Set btnMenuExtendet(ICNT).PressedImage = picBuff.Image
    
    btnMenuExtendet(ICNT).Left = 2
    If ICNT = 1 Then
        btnMenuExtendet(ICNT).Top = 2 + (ICNT - 1) * btnMenuExtendet(ICNT).Height
    Else
        btnMenuExtendet(ICNT).Top = 5 + (ICNT - 1) * btnMenuExtendet(ICNT).Height
    
    End If
    btnMenuExtendet(ICNT).Visible = True
Next ICNT
'set up banners
picBanExtendet(0).Top = btnMenuExtendet(1).Top + btnMenuExtendet(1).Height
picBanExtendet(0).Left = 2
picBanExtendet(0).Width = btnMenuExtendet(1).Width + 1
picBanExtendet(0).Height = 3

firsColor = vb3DHighlight
SecondColor = AlphaBlend(&H8000000F, vb3DDKShadow, 240)

picBanExtendet(0).BackColor = PicMnuExtendet.BackColor

For IcNT2 = 1 To btnMenuExtendet(1).Height
    picBanExtendet(0).ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * IcNT2 / btnMenuExtendet(1).Height))
    picBanExtendet(0).Line (IcNT2 - 1, 0)-(IcNT2 - 1, 3)
    picBanExtendet(0).Refresh
Next IcNT2

picBanExtendet(0).ForeColor = AlphaBlend(SecondColor, shpBorder.BorderColor, 128)

picBanExtendet(0).Line (btnMenuExtendet(1).Height + 8, 1)-(picBanExtendet(0).Width, 1)

picBanExtendet(0).Visible = True
'SETS THE WHOLE MENU
Line1.BorderColor = PicMnuExtendet.BackColor
Line1.X1 = 1
Line1.X2 = picDetails.Width - 3
Line1.Y1 = 0
Line1.Y2 = 0
Line1.ZOrder

PicMnuExtendet.Width = (picBuff.Width + 8) * Screen.TwipsPerPixelX
PicMnuExtendet.Height = (picBuff.Height * (btnMenuExtendet.Count - 1) + 11) * Screen.TwipsPerPixelY

shpBorder.BorderColor = AlphaBlend(SecondColor, vb3DDKShadow, 150)
shpBorder.Left = 0
shpBorder.Top = 0
shpBorder.Width = PicMnuExtendet.Width / Screen.TwipsPerPixelX - 4
shpBorder.Height = PicMnuExtendet.Height / Screen.TwipsPerPixelY - 4

picBuff3.Height = PicMnuExtendet.Height / Screen.TwipsPerPixelY - 4
picBuff3.Width = 4
picBuff3.Cls

DeskHdc = GetDC(0)

ret = BitBlt(picBuff3.hDC, 0, 0, picBuff3.Width, picBuff3.Height, DeskHdc, PicMnuExtendet.Left / Screen.TwipsPerPixelX + PicMnuExtendet.Width / Screen.TwipsPerPixelY - 4, PicMnuExtendet.Top / Screen.TwipsPerPixelY + 4, SRCCOPY)
ret = BitBlt(PicMnuExtendet.hDC, PicMnuExtendet.Width / Screen.TwipsPerPixelY - 4, 0, picBuff3.Width, picBuff3.Height, DeskHdc, PicMnuExtendet.Left / Screen.TwipsPerPixelX + PicMnuExtendet.Width / Screen.TwipsPerPixelY - 4, PicMnuExtendet.Top / Screen.TwipsPerPixelY, SRCCOPY)

picBuff3.Picture = picBuff3.Image

Set cc = DrawShadow(picBuff3.Picture, True)
cc.BitBlt PicMnuExtendet.hDC, PicMnuExtendet.Width / Screen.TwipsPerPixelY - 4, 4, picBuff3.Width, picBuff3.Height, 0, 0

picBuff3.Height = 4
picBuff3.Width = PicMnuExtendet.Width / Screen.TwipsPerPixelX - 3

ret = BitBlt(picBuff3.hDC, 0, 0, picBuff3.Width, picBuff3.Height, DeskHdc, PicMnuExtendet.Left / Screen.TwipsPerPixelX, PicMnuExtendet.Top / Screen.TwipsPerPixelY + PicMnuExtendet.Height / Screen.TwipsPerPixelY - 4, SRCCOPY)

picBuff3.Picture = picBuff3.Image

Set cc = DrawShadow(picBuff3.Picture, False)
cc.BitBlt PicMnuExtendet.hDC, 0, PicMnuExtendet.Height / Screen.TwipsPerPixelY - 4, picBuff3.Width, picBuff3.Height, 0, 0

picShdw.Height = (btn.Height + 2) * Screen.TwipsPerPixelY
picShdw.Width = 4 * Screen.TwipsPerPixelX

picBuff3.Width = picShdw.Width / Screen.TwipsPerPixelX
picBuff3.Height = picShdw.Height / Screen.TwipsPerPixelX

picShdw.Cls
picShdw.Picture = LoadPicture("")
ret = BitBlt(picShdw.hDC, 0, 0, picShdw.Width, picShdw.Height, DeskHdc, PicMnuExtendet.Left / Screen.TwipsPerPixelX + btn.Width, PicMnuExtendet.Top / Screen.TwipsPerPixelY - btn.Height + 2, SRCCOPY)
ret = ReleaseDC(0&, DeskHdc)

picShdw.Top = PicMnuExtendet.Top - (btn.Height - 2) * Screen.TwipsPerPixelY
picShdw.Left = PicMnuExtendet.Left + (picDetails.Width - 2) * Screen.TwipsPerPixelX

picShdw.Picture = picShdw.Image

Set cc = DrawShadow(picShdw.Picture, True)
cc.BitBlt picShdw.hDC, 0, 0, picShdw.Width, picShdw.Height, 0, 0

End Sub

Private Sub SetUpMnu(index As Long)
Dim ICNT As Integer
Dim cc As cMemDC
Dim caption As String
Dim Tag As String
Dim IconIndex As Integer
Dim bTop As Integer
Dim DrawSelected As Boolean
Dim MyDocumentsFolder As String
MyDocumentsFolder = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
If index > ListCount Then index = 0
PicMnu.Tag = index
'SETS BUTTONS
For ICNT = 1 To btnMenu.Count - 1
    Unload btnMenu(ICNT)
Next ICNT
Dim K  As Integer
K = 6
Dim tWidth As Integer
tWidth = picMnuText.TextWidth(MNU_NEWFOLDER)
If tWidth < picMnuText.TextWidth(MNU_REFRESH) Then tWidth = picMnuText.TextWidth(MNU_REFRESH)
If tWidth < picMnuText.TextWidth(MNUVIEW_ETENDETMODE) Then tWidth = picMnuText.TextWidth(MNUVIEW_ETENDETMODE)
If tWidth < picMnuText.TextWidth(MNUVIEW_LIST_SMALL) Then tWidth = picMnuText.TextWidth(MNUVIEW_LIST_SMALL)
If tWidth < picMnuText.TextWidth(MNUVIEW_LIST_LARGE) Then tWidth = picMnuText.TextWidth(MNUVIEW_LIST_LARGE)
If tWidth < picMnuText.TextWidth(MNUVIEW_ICONS) Then tWidth = picMnuText.TextWidth(MNUVIEW_ICONS)
'If tWidth < picMnuText.TextWidth(MNU_COPY) Then tWidth = picMnuText.TextWidth(MNU_COPY)
'If tWidth < picMnuText.TextWidth(MNU_CUT) Then tWidth = picMnuText.TextWidth(MNU_CUT)
'If tWidth < picMnuText.TextWidth(MNU_PASTE) Then tWidth = picMnuText.TextWidth(MNU_PASTE)

If index > 0 Then
    K = K + 4
    picMnuText.Font.Bold = True
    If tWidth < picMnuText.TextWidth(MNU_SELECT) Then tWidth = picMnuText.TextWidth(MNU_SELECT)
    picMnuText.Font.Bold = False
    If tWidth < picMnuText.TextWidth(MNU_RENAME) Then tWidth = picMnuText.TextWidth(MNU_RENAME)
    If tWidth < picMnuText.TextWidth(MNU_PROPERTYS) Then tWidth = picMnuText.TextWidth(MNU_PROPERTYS)
    If tWidth < picMnuText.TextWidth(MNU_DELETE) Then tWidth = picMnuText.TextWidth(MNU_DELETE)
End If

tWidth = tWidth + 5

Dim DrawDisabled As Boolean

For ICNT = 1 To K
    picMnuText.Font.Bold = False
    DrawDisabled = False
    DrawSelected = False
    If index > 0 Then
        If ICNT = 1 Then
            caption = MNU_SELECT
            Tag = "#SEL"
            IconIndex = 0
            bTop = 0
            picMnuText.Font.Bold = True
        ElseIf ICNT = 2 Then
            caption = MNU_RENAME
            Tag = "#REN"
            IconIndex = 0
            bTop = 3
            If Path = "#MYCOMPUTER" Or iFile(PicMnu.Tag).FullName = "#MYCOMPUTER" Or iFile(PicMnu.Tag).FullName = MyDocumentsFolder Then
                DrawDisabled = True
            End If
        ElseIf ICNT = 3 Then
            caption = MNU_DELETE
            Tag = "#DEL"
            IconIndex = 3
            bTop = 3
            If Path = "#MYCOMPUTER" Or iFile(PicMnu.Tag).FullName = "#MYCOMPUTER" Or iFile(PicMnu.Tag).FullName = MyDocumentsFolder Then
                DrawDisabled = True
            End If
        ElseIf ICNT = 4 Then
            caption = MNU_REFRESH
            Tag = "#REF"
            IconIndex = 0
            bTop = 6
        ElseIf ICNT = 5 Then
            caption = MNU_NEWFOLDER
            Tag = "#NEW"
            IconIndex = 1
            bTop = 6
            If Path = "#MYCOMPUTER" Then
                DrawDisabled = True
            End If
        ElseIf ICNT = 6 Then
            caption = MNUVIEW_ETENDETMODE
            Tag = "#EXT"
            IconIndex = 2
            bTop = 9
            mnuIcon(2).Picture = MnuViewIcon(0)
            If dView = 0 Then DrawSelected = True
        ElseIf ICNT = 7 Then
            caption = MNUVIEW_LIST_SMALL
            Tag = "#LSM"
            IconIndex = 2
            bTop = 9
            mnuIcon(2).Picture = MnuViewIcon(1)
            If dView = 1 Then DrawSelected = True
        ElseIf ICNT = 8 Then
            caption = MNUVIEW_LIST_LARGE
            Tag = "#LSL"
            IconIndex = 2
            bTop = 9
            mnuIcon(2).Picture = MnuViewIcon(2)
            If dView = 2 Then DrawSelected = True
        ElseIf ICNT = 9 Then
            caption = MNUVIEW_ICONS
            Tag = "#ICO"
            IconIndex = 2
            bTop = 9
            mnuIcon(2).Picture = MnuViewIcon(3)
            If dView = 3 Then DrawSelected = True
        ElseIf ICNT = 10 Then
            caption = MNU_PROPERTYS
            Tag = "#PRO"
            IconIndex = 0
            bTop = 12
        End If
    Else
        If ICNT = 1 Then
            caption = MNU_REFRESH
            Tag = "#REF"
            IconIndex = 0
            bTop = 0
        ElseIf ICNT = 2 Then
            caption = MNU_NEWFOLDER
            Tag = "#NEW"
            IconIndex = 1
            bTop = 0
            If Path = "#MYCOMPUTER" Then
                DrawDisabled = True
            End If
        ElseIf ICNT = 3 Then
            caption = MNUVIEW_ETENDETMODE
            Tag = "#EXT"
            IconIndex = 2
            bTop = 3
            mnuIcon(2).Picture = MnuViewIcon(0)
            If dView = 0 Then DrawSelected = True
        ElseIf ICNT = 4 Then
            caption = MNUVIEW_LIST_SMALL
            Tag = "#LSM"
            IconIndex = 2
            bTop = 3
            mnuIcon(2).Picture = MnuViewIcon(1)
            If dView = 1 Then DrawSelected = True
        ElseIf ICNT = 5 Then
            caption = MNUVIEW_LIST_LARGE
            Tag = "#LSL"
            IconIndex = 2
            bTop = 3
            mnuIcon(2).Picture = MnuViewIcon(2)
            If dView = 2 Then DrawSelected = True
        ElseIf ICNT = 6 Then
            caption = MNUVIEW_ICONS
            Tag = "#ICO"
            IconIndex = 2
            bTop = 3
            mnuIcon(2).Picture = MnuViewIcon(3)
            If dView = 3 Then DrawSelected = True
        End If
    End If
    Load btnMenu(ICNT)
    btnMenu(ICNT).Enabled = Not DrawDisabled
    btnMenu(ICNT).Tag = Tag
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
    SecondColor = AlphaBlend(vb3DFace, vb3DDKShadow, 240)

    Dim IcNT2 As Integer
    picBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 40)
    For IcNT2 = 1 To picBuff3.Width
        picBuff3.ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * IcNT2 / picBuff3.Width))
        picBuff3.Line (IcNT2 - 1, 0)-(IcNT2 - 1, picBuff3.Height)
        picBuff3.Refresh
    Next IcNT2
    
    If DrawSelected = True Then
        
        If DrawDisabled = True Then
        Else
            picBuff3.ForeColor = AlphaBlend(vbHighlight, vbWindowBackground, 150)
        End If
        
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
    End If
    
    BitBlt picBuff.hDC, 0, 0, picBuff3.Width, picBuff3.Height, picBuff3.hDC, 0, 0, SRCCOPY
    
    picBuff3.Cls
    picBuff3.Picture = mnuIcon(IconIndex)

    picBuff3.Cls
    picBuff3.Picture = LoadPicture("")
    picMnuText.Width = tWidth
    picBuff3.BackColor = picBuff.BackColor
    
    picMnuText.Height = picMnuText.TextHeight("Žg")
    picMnuText.BackColor = picBuff.BackColor
    
    If DrawDisabled = True Then
        picMnuText.ForeColor = AlphaBlend(vb3DDKShadow, vb3DHighlight, 70)
    Else
        picMnuText.ForeColor = &H80000007
    End If
    
    picMnuText.Cls
    
    picMnuText.Print caption

    
    BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picMnuText.Height) / 2, picMnuText.Width, picMnuText.Height, picMnuText.hDC, 0, 0, SRCCOPY
    
    picBuff3.Picture = mnuIcon(IconIndex)


    If DrawDisabled = True Then
        If DrawSelected = True Then
            Set cc = DisabledPicture(picBuff3.Picture, AlphaBlend(vbHighlight, vbWindowBackground, 40), True)
            cc.BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, 0, 0
        Else
            Set cc = DisabledPicture(picBuff3.Picture, picBuff3.BackColor, False)
            cc.BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, 0, 0
        End If
    Else
        If DrawSelected = True Then
            Set cc = DimBitmap2(picBuff3.Picture, AlphaBlend(vbHighlight, vbWindowBackground, 40), True)
            cc.BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, 0, 0
        Else
            Set cc = DimBitmap2(picBuff3.Picture, picBuff3.BackColor, False)
            cc.BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, 0, 0
        End If
    End If
    
    Set btnMenu(ICNT).NormalImage = picBuff.Image
    
    picBuff.Cls
    
    If DrawDisabled = True Then
        picBuff.ForeColor = AlphaBlend(vb3DDKShadow, vb3DHighlight, 70)
        picBuff.BackColor = AlphaBlend(picBuff.ForeColor, vb3DHighlight, 70)
    Else
        picBuff.ForeColor = vbHighlight
        picBuff.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
    End If
    
    picBuff3.Picture = LoadPicture("")
    
    If DrawSelected = True Then
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
    picBuff3.BackColor = picBuff.BackColor
    picBuff3.Picture = mnuIcon(IconIndex)
    
    If DrawDisabled = True Then
        KCY = True
        If DrawSelected = True Then
            Set cc = DisabledPicture(picBuff3.Picture, AlphaBlend(vbHighlight, vbWindowBackground, 40), True)
            cc.BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, 0, 0
        Else
            Set cc = DisabledPicture(picBuff3.Picture, picBuff3.BackColor, False)
            cc.BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, 0, 0
        End If
        KCY = False
    Else
        BitBlt picBuff.hDC, (picBuff.Height - 16) / 2, (picBuff.Height - 16) / 2, 18, 18, picBuff3.hDC, 0, 0, SRCCOPY
    End If
    picMnuText.Cls
    picBuff3.Picture = LoadPicture("")
    picMnuText.Width = tWidth
    picMnuText.BackColor = picBuff.BackColor
    
    If DrawDisabled = True Then
        picMnuText.ForeColor = AlphaBlend(vb3DDKShadow, vb3DHighlight, 70)
    Else
        picMnuText.ForeColor = &H80000007
    End If
    
    picMnuText.Print caption

    BitBlt picBuff.hDC, picBuff.Height + 3, (picBuff.Height - picMnuText.Height) / 2, picMnuText.Width, picMnuText.Height, picMnuText.hDC, 0, 0, SRCCOPY
    
   ' If Tag = "#VIW" Then
   '     picBuff.ForeColor = &H80000007
   '     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2)-(picBuff.Width - 7, picBuff.Height / 2)
   '     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2 - 1)-(picBuff.Width - 8, picBuff.Height / 2 - 1)
   '     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2 + 1)-(picBuff.Width - 8, picBuff.Height / 2 + 1)
   ''     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2 - 2)-(picBuff.Width - 9, picBuff.Height / 2 - 2)
   '     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2 + 2)-(picBuff.Width - 9, picBuff.Height / 2 + 2)
   '     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2 - 3)-(picBuff.Width - 10, picBuff.Height / 2 - 3)
   '     picBuff.Line (picBuff.Width - 11, picBuff.Height / 2 + 3)-(picBuff.Width - 10, picBuff.Height / 2 + 3)
   ' End If
    
    Set btnMenu(ICNT).FocusedImage = picBuff.Image
    Set btnMenu(ICNT).PressedImage = picBuff.Image
    
    btnMenu(ICNT).Left = 2

    btnMenu(ICNT).Top = 2 + (ICNT - 1) * btnMenu(ICNT).Height + bTop

    btnMenu(ICNT).Visible = True
Next ICNT
'set up banners
For ICNT = 1 To picBan.Count - 1
    Unload picBan(ICNT)
Next ICNT

If index > 0 Then
    For ICNT = 1 To 4
    Load picBan(ICNT)
    If ICNT = 1 Then
        picBan(ICNT).Top = btnMenu(1).Top + btnMenu(1).Height
    ElseIf ICNT = 2 Then
        picBan(ICNT).Top = btnMenu(3).Top + btnMenu(3).Height
    
    ElseIf ICNT = 3 Then
        picBan(ICNT).Top = btnMenu(5).Top + btnMenu(5).Height
    
    ElseIf ICNT = 4 Then
        picBan(ICNT).Top = btnMenu(9).Top + btnMenu(9).Height
    End If
    
    picBan(ICNT).Left = 2
    picBan(ICNT).Width = btnMenu(2).Width + 1
    picBan(ICNT).Height = 3
    
    firsColor = vb3DHighlight
    SecondColor = AlphaBlend(&H8000000F, vb3DDKShadow, 240)
    
    picBan(ICNT).BackColor = PicMnu.BackColor
    
    For IcNT2 = 1 To btnMenu(1).Height
        picBan(ICNT).ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * IcNT2 / btnMenu(1).Height))
        picBan(ICNT).Line (IcNT2 - 1, 0)-(IcNT2 - 1, 3)
        picBan(ICNT).Refresh
    Next IcNT2
    
    picBan(ICNT).ForeColor = AlphaBlend(SecondColor, shp.BorderColor, 128)
    
    picBan(ICNT).Line (btnMenu(1).Height + 8, 1)-(picBan(ICNT).Width, 1)
    
    picBan(ICNT).Visible = True
    Next ICNT
Else
    For ICNT = 1 To 1
    Load picBan(ICNT)
    If ICNT = 1 Then
        picBan(ICNT).Top = btnMenu(2).Top + btnMenu(2).Height
    End If
    
    picBan(ICNT).Left = 2
    picBan(ICNT).Width = btnMenu(2).Width + 1
    picBan(ICNT).Height = 3
    
    firsColor = vb3DHighlight
    SecondColor = AlphaBlend(&H8000000F, vb3DDKShadow, 240)
    
    picBan(ICNT).BackColor = PicMnu.BackColor
    
    For IcNT2 = 1 To btnMenu(1).Height
        picBan(ICNT).ForeColor = AlphaBlend(SecondColor, firsColor, Int(255 * IcNT2 / btnMenu(1).Height))
        picBan(ICNT).Line (IcNT2 - 1, 0)-(IcNT2 - 1, 3)
        picBan(ICNT).Refresh
    Next IcNT2
    
    picBan(ICNT).ForeColor = AlphaBlend(SecondColor, shp.BorderColor, 128)
    
    picBan(ICNT).Line (btnMenu(1).Height + 8, 1)-(picBan(ICNT).Width, 1)
    
    picBan(ICNT).Visible = True
    Next ICNT
End If
'SETS THE WHOLE MENU

PicMnu.Width = (picBuff.Width + 8) * Screen.TwipsPerPixelX
PicMnu.Height = (picBuff.Height * (btnMenu.Count - 1) + bTop + 8) * Screen.TwipsPerPixelY

If PicMnu.Top + PicMnu.Height > Screen.Height - 32 * Screen.TwipsPerPixelY Then
    If PicMnu.Top > PicMnu.Height + 5 * Screen.TwipsPerPixelY Then
        PicMnu.Top = PicMnu.Top - PicMnu.Height + 5 * Screen.TwipsPerPixelY
    Else
        PicMnu.Top = Screen.Height - 32 * Screen.TwipsPerPixelY - PicMnu.Height
    End If
End If

my = PicMnu.Top

shp.BorderColor = AlphaBlend(SecondColor, vb3DDKShadow, 150)
shp.Left = 0
shp.Top = 0
shp.Width = PicMnu.Width / Screen.TwipsPerPixelX - 4
shp.Height = PicMnu.Height / Screen.TwipsPerPixelY - 4

picBuff3.Height = PicMnu.Height / Screen.TwipsPerPixelY - 4
picBuff3.Width = 4
picBuff3.Cls

DeskHdc = GetDC(0)

ret = BitBlt(picBuff3.hDC, 0, 0, picBuff3.Width, picBuff3.Height, DeskHdc, PicMnu.Left / Screen.TwipsPerPixelX + PicMnu.Width / Screen.TwipsPerPixelY - 4, PicMnu.Top / Screen.TwipsPerPixelY + 4, SRCCOPY)
ret = BitBlt(PicMnu.hDC, PicMnu.Width / Screen.TwipsPerPixelY - 4, 0, picBuff3.Width, picBuff3.Height, DeskHdc, PicMnu.Left / Screen.TwipsPerPixelX + PicMnu.Width / Screen.TwipsPerPixelY - 4, PicMnu.Top / Screen.TwipsPerPixelY, SRCCOPY)

picBuff3.Picture = picBuff3.Image

Set cc = DrawShadow(picBuff3.Picture, True)
cc.BitBlt PicMnu.hDC, PicMnu.Width / Screen.TwipsPerPixelY - 4, 4, picBuff3.Width, picBuff3.Height, 0, 0

picBuff3.Height = 4
picBuff3.Width = PicMnu.Width / Screen.TwipsPerPixelX - 3

ret = BitBlt(picBuff3.hDC, 0, 0, picBuff3.Width, picBuff3.Height, DeskHdc, PicMnu.Left / Screen.TwipsPerPixelX, PicMnu.Top / Screen.TwipsPerPixelY + PicMnu.Height / Screen.TwipsPerPixelY - 4, SRCCOPY)
ret = ReleaseDC(0&, DeskHdc)

picBuff3.Picture = picBuff3.Image

Set cc = DrawShadow(picBuff3.Picture, False)
cc.BitBlt PicMnu.hDC, 0, PicMnu.Height / Screen.TwipsPerPixelY - 4, picBuff3.Width, picBuff3.Height, 0, 0

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
                    If KCY = False Then .SetPixel lI, lJ, AlphaBlend(AlphaBlend(frmDialog.BackColor, vb3DDKShadow, 240), vb3DHighlight, Int(255 * (5 + lI) / 25))
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
                    .SetPixel lI, lJ, AlphaBlend(vb3DFace, .GetPixel(lI, lJ), 70)
                Else
                    If IsSelected = True Then
                        .SetPixel lI, lJ, MaskColor
                    Else
                        .SetPixel lI, lJ, AlphaBlend(AlphaBlend(vb3DFace, vb3DDKShadow, 240), vb3DHighlight, Int(255 * (5 + lI) / 25))
                    End If
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
    Set cc = picBuff3
    
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

Private Sub btn_Click()
If PicMnuExtendet.Visible = True Then
    picBuff3.BackColor = vb3DHighlight
    picBuff3.ForeColor = shpBorder.BorderColor
    
    If iFile(btn.Tag).ShowDetails = False Then
        picBuff3.Height = SelHeight - 4
        picBuff3.Width = picDetails.Width - 2
        
        picDetails.BackColor = picBuff3.BackColor
        BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
        
        picBuff3.Line (0, 0)-(picBuff3.Width, 0)
        picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
        picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
        picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)
        Set btn.NormalImage = picBuff3.Image
        Set btn.FocusedImage = picBuff3.Image
    Else
        picBuff3.Cls
        picBuff3.Height = selDetHeight - 4
        picBuff3.Width = picDetails.Width - 2
        
        picHideDetails.BackColor = picBuff3.BackColor
        BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
        
        picBuff3.Line (0, 0)-(picBuff3.Width, 0)
        picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
        picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
        picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)
        
        Set btn.NormalImage = picBuff3.Image
        Set btn.FocusedImage = picBuff3.Image
        
        picBuff3.ForeColor = vbHighlight
        picBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)
    
        picHideDetails.BackColor = picBuff3.BackColor
        BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picHideDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
        
        picBuff3.Line (0, 0)-(picBuff3.Width, 0)
        picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
        picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
        picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)
    
        Set btn.PressedImage = picBuff3.Image
    End If

    PicMnuExtendet.Top = my
    PicMnuExtendet.Left = mx
    picShdw.Visible = True
    picShdw.ZOrder
    PicMnuExtendet.ZOrder
    SetCapture PicMnuExtendet.hwnd
End If

End Sub

Private Sub btn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
Dim Rec As RECT
Dim ICNT As Long

If dView = 0 Then
    ICNT = btn.Tag
    If Button = vbRightButton Then
    
        If iFile(ICNT).ShowDetails = True And GetFileTop(ICNT) > picOzadje.Height - selDetHeight - iTop Then
            iTop = picOzadje.Height - selDetHeight - GetFileTop(ICNT)
            UserControl_Resize
        End If

        GetWindowRect UserControl.hwnd, Rec
        
        mx = (Rec.Left + 3) * Screen.TwipsPerPixelX
        If iFile(ICNT).ShowDetails = True Then
            my = (Rec.Top + iTop + GetFileTop(ICNT) + selDetHeight - 2) * Screen.TwipsPerPixelY
        Else
            my = (Rec.Top + iTop + GetFileTop(ICNT) + SelHeight - 2) * Screen.TwipsPerPixelY
        End If
        
        PicMnuExtendet.Top = my
        PicMnuExtendet.Left = mx
        
        SetUpMnuExtendet ICNT
        
        PicMnuExtendet.Top = -PicMnuExtendet.Height
        
        PicMnuExtendet.Visible = True
        Exit Sub
    End If

End If

End Sub

Private Sub btn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
picOzadje_MouseDown Button, Shift, X + btn.Left, Y + btn.Top

picBuff3.ForeColor = vbHighlight
picBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
picBuff3.Cls
picBuff3.Picture = LoadPicture("")
If iFile(btn.Tag).ShowDetails = False Then
     picBuff3.Height = SelHeight - 4
     picBuff3.Width = picDetails.Width - 2
     
     picDetails.BackColor = picBuff3.BackColor
     BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
     
     picBuff3.Line (0, 0)-(picBuff3.Width, 0)
     picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
     picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
     picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)

     Set btn.NormalImage = picBuff3.Image
     Set btn.FocusedImage = picBuff3.Image
     
     picBuff3.ForeColor = vbHighlight
     picBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

     picDetails.BackColor = picBuff3.BackColor
     BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
     
     picBuff3.Line (0, 0)-(picBuff3.Width, 0)
     picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
     picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
     picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)

     Set btn.PressedImage = picBuff3.Image
     
Else
     picBuff3.Cls
     picBuff3.Height = selDetHeight - 4
     picBuff3.Width = picDetails.Width - 2
     
     picHideDetails.BackColor = picBuff3.BackColor
     BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
     
     picBuff3.Line (0, 0)-(picBuff3.Width, 0)
     picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
     picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
     picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)
     
     Set btn.NormalImage = picBuff3.Image
     Set btn.FocusedImage = picBuff3.Image
     
     picBuff3.ForeColor = vbHighlight
     picBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

     picHideDetails.BackColor = picBuff3.BackColor
     BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picHideDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
     
     picBuff3.Line (0, 0)-(picBuff3.Width, 0)
     picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
     picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
     picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)

     Set btn.PressedImage = picBuff3.Image
 End If
 
 If btn.Top <> GetFileTop(btn.Tag) + iTop + 2 Then
    btn.Top = GetFileTop(btn.Tag) + iTop + 2
 End If
 
End Sub

Private Sub btn_MouseOut()
btn.Visible = False

End Sub

Private Sub btnDOWN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Button = vbLeftButton Then
    tmrScr.Tag = "D"
    tmrScr.Interval = 1500
    tmrScr_Timer
    tmrScr.Enabled = True
End If

End Sub

Private Sub btnDOWN_TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrScr.Enabled = False

End Sub

Private Sub btnLEFT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Button = vbLeftButton Then
    tmrScr.Tag = "L"
    tmrScr.Interval = 1500
    tmrScr_Timer
    tmrScr.Enabled = True
End If

End Sub


Private Sub btnLEFT_TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrScr.Enabled = False

End Sub

Private Sub btnRIGHT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Button = vbLeftButton Then
    tmrScr.Tag = "R"
    tmrScr.Interval = 1500
    tmrScr_Timer
    tmrScr.Enabled = True
End If

End Sub


Private Sub btnRIGHT_TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrScr.Enabled = False

End Sub


Private Sub btnUP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
If Button = vbLeftButton Then
    tmrScr.Tag = "U"
    tmrScr.Interval = 1500
    tmrScr_Timer
    tmrScr.Enabled = True
End If

End Sub

Private Sub btnUP_TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrScr.Enabled = False

End Sub

Private Sub PicMnu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Integer
Dim gen As Boolean
Dim IcNT2 As Long
If X < 0 Or Y < 0 Or X > PicMnu.Width Or Y > PicMnu.Height / Screen.TwipsPerPixelY Then

Else
    For ICNT = 1 To btnMenu.Count - 1
        If Button = vbLeftButton And X >= 0 And X < PicMnu.Width / Screen.TwipsPerPixelX - 4 And Y >= btnMenu(ICNT).Top And Y < btnMenu(ICNT).Top + btnMenu(ICNT).Height Then
            If btnMenu(ICNT).cEnabled = False Then Exit For
            
            If btnMenu(ICNT).Tag = "#NEW" Then
                RaiseEvent MenuNewFolderClick
            ElseIf btnMenu(ICNT).Tag = "#SEL" Then
                If iFile(PicMnu.Tag).Extension = "#FOLDER" Then
                    RaiseEvent DirSelect(iFile(PicMnu.Tag).FullName)
                    Path = iFile(PicMnu.Tag).FullName
                Else
                    If bSelected > 0 Then RaiseEvent FileSelect(GetSelectedFiles)
                End If
            ElseIf btnMenu(ICNT).Tag = "#EXT" Then
                View = 0
            ElseIf btnMenu(ICNT).Tag = "#LSM" Then
                View = 1
            ElseIf btnMenu(ICNT).Tag = "#LSL" Then
                View = 2
            ElseIf btnMenu(ICNT).Tag = "#ICO" Then
                View = 3
            ElseIf btnMenu(ICNT).Tag = "#REF" Then
                Refresh
            ElseIf btnMenu(ICNT).Tag = "#PRO" Then
                    Dim SEI As SHELLEXECUTEINFO
                    With SEI
                        .cbSize = Len(SEI)
                        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
                            SEE_MASK_INVOKEIDLIST Or _
                            SEE_MASK_FLAG_NO_UI
                        .hwnd = UserControl.hwnd
                        .lpVerb = "properties"
                        .lpFile = iFile(PicMnu.Tag).FullName
                        .lpParameters = vbNullChar
                        .lpDirectory = vbNullChar
                        .nShow = 0
                        .hInstApp = 0
                        .lpIDList = 0
                    End With
                    Call ShellExecuteEx(SEI)
            ElseIf btnMenu(ICNT).Tag = "#DEL" Then
                Dim ref As Boolean
                For IcNT2 = 1 To ListCount
                    If iFile(IcNT2).Selected = True Then
                        If RemoveFile(iFile(IcNT2).FullName, RecycleFile) = True Then
                            ref = True
                        End If
                    End If
                Next IcNT2
                If ref = True Then Refresh
            ElseIf btnMenu(ICNT).Tag = "#REN" Then
                Dim KY As Long
                Dim KX As Long
                txtRename.Font.Size = 8
                txtRename.Font.Bold = False
                If dView = 0 Then
                    If iFile(PicMnu.Tag).ShowDetails = False Then
                        KX = 30
                        KY = GetFileTop(PicMnu.Tag) + 1 + iTop
                        txtRename.Font = picBuffText.Font
                        txtRename.Tag = PicMnu.Tag
                        PicRename.Move KX, KY, picOzadje.Width - KX - 1, SelHeight - 2
                        
                        txtRename.BackColor = vbWhite
                        txtRename.ForeColor = SelForeColor
                        PicRename.BackColor = vbWhite
                        txtRename.Move 1, 0, PicRename.Width - 1, PicRename.Height
                    Else
                        KX = 56
                        KY = GetFileTop(PicMnu.Tag) + 1 + iTop
                        txtRename.Font = picBuffText.Font
                        txtRename.Font.Size = 9
                        txtRename.Font.Bold = True
                        txtRename.Tag = PicMnu.Tag
                        PicRename.Move KX, KY, picOzadje.Width - KX - 1, SelHeight - 3
                        
                        txtRename.BackColor = vbWhite
                        txtRename.ForeColor = SelForeColor
                        PicRename.BackColor = vbWhite
                        txtRename.Move 1, 0, PicRename.Width - 1, PicRename.Height
                    End If
                ElseIf dView = 1 Then
                    If ColumnCount > 1 Then
                        KX = Int((PicMnu.Tag - 1) / (RowsInColumnCount))
                        KY = PicMnu.Tag - (KX) * RowsInColumnCount - 1
                    Else
                        KY = PicMnu.Tag - 1
                        KX = 0
                    End If
                    'PicRename.Top =
                    KX = KX * ColumnWidth + 18 + iLeft
                    KY = KY * picBackground.Height / 2 + 1
                    txtRename.Font = picBuffText.Font
                    txtRename.Tag = PicMnu.Tag
                    PicRename.Move KX, KY, ColumnWidth - 19, picBackground.Height / 2 - 2
                    txtRename.BackColor = vbWhite
                    txtRename.ForeColor = SelForeColor
                    PicRename.BackColor = vbWhite
                    txtRename.Move 0, 0, PicRename.Width, PicRename.Height
                
                ElseIf dView = 2 Then
                    If ColumnCount > 1 Then
                        KX = Int((PicMnu.Tag - 1) / (RowsInColumnCount))
                        KY = PicMnu.Tag - (KX) * RowsInColumnCount - 1
                    Else
                        KY = PicMnu.Tag - 1
                        KX = 0
                    End If
                    
                    'PicRename.Top =
                    KX = KX * ColumnWidth + 34 + iLeft
                    KY = KY * picBackground.Height / 2 + 1
                    txtRename.Font = picBuffText.Font
                    txtRename.Tag = PicMnu.Tag
                    PicRename.Move KX, KY, ColumnWidth - 35, picBackground.Height / 2 - 2
                    txtRename.BackColor = vbWhite
                    txtRename.ForeColor = SelForeColor
                    PicRename.BackColor = vbWhite
                    txtRename.Move 0, (PicRename.Height - picBuffText.Height) / 2, PicRename.Width, picBuffText.Height
                
                ElseIf dView = 3 Then
                    If ColumnCount > 1 Then
                        KY = Int((PicMnu.Tag - 1) / (RowsInColumnCount))
                        KX = PicMnu.Tag - (KY) * RowsInColumnCount - 1
                    Else
                        KX = PicMnu.Tag - 1
                        KY = 0
                    End If
                    
                    KX = KX * 96 + 1
                    KY = KY * picBackground.Height / 2 + 35 + iTop
                    txtRename.Font = picBuffText.Font
                    txtRename.Tag = PicMnu.Tag
                    PicRename.Move KX, KY, 94, 29
                    
                    txtRename.BackColor = vbWhite
                    txtRename.ForeColor = SelForeColor
                    PicRename.BackColor = vbWhite
                    txtRename.Move 1, 0, PicRename.Width - 1, PicRename.Height
                End If
                If dView = 3 Then txtRename.Alignment = 2 Else txtRename.Alignment = 0
                txtRename.Text = iFile(PicMnu.Tag).name
                PicRename.Visible = True
                txtRename.SetFocus
            End If
        End If
    Next ICNT
End If

ReleaseCapture
PicMnu.Visible = False

'If gen = True Then UserControl_Resize

End Sub

Private Function RemoveFile(FileName As String, Action As RemoveMethod) As Boolean
'* Copyright (c) 2001 Nicholas Skapura * "Advance File Access" - this code was taken from him - found on PSC...
    Dim FileOperation As SHFILEOPSTRUCT
    Dim tmpReturn As Long
    On Error GoTo RemoveFile_Err
    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = FileName
        If Action = RecycleFile Then
            .fFlags = FOF_ALLOWUNDO + FOF_CREATEPROGRESSDLG
        Else
            .fFlags = FO_DELETE + FOF_CREATEPROGRESSDLG
        End If
    End With
    tmpReturn = SHFileOperation(FileOperation)
    If tmpReturn <> 0 Then
        RemoveFile = False
    Else
        RemoveFile = True
    End If
    Exit Function
RemoveFile_Err:
    RemoveFile = False
    
End Function
Private Sub PicMnu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Integer

For ICNT = 1 To btnMenu.Count - 1
    If X >= 0 And X < PicMnu.Width / Screen.TwipsPerPixelX - 4 And Y >= btnMenu(ICNT).Top And Y < btnMenu(ICNT).Top + btnMenu(ICNT).Height Then
        btnMenu(ICNT).SetState "H"
    Else
        btnMenu(ICNT).SetState "N"
    End If
Next ICNT

End Sub

Private Sub PicMnuExtendet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Integer
Dim gen As Boolean
Dim IcNT2 As Long
If X < 0 Or Y < 0 Or X > PicMnuExtendet.Width Or Y > PicMnuExtendet.Height Then

Else
    For ICNT = 1 To btnMenuExtendet.Count - 1
        If Button = vbLeftButton And X >= 0 And X < PicMnuExtendet.Width / Screen.TwipsPerPixelX - 4 And Y >= btnMenuExtendet(ICNT).Top And Y < btnMenuExtendet(ICNT).Top + btnMenuExtendet(ICNT).Height Then
            If ICNT = 1 Then
                iFile(btnMenuExtendet(ICNT).Tag).ShowDetails = Not iFile(btnMenuExtendet(ICNT).Tag).ShowDetails
                gen = True
                If GetFileTop(btnMenuExtendet(ICNT).Tag) > picOzadje.Height - selDetHeight Then
                    iTop = picOzadje.Height - selDetHeight - GetFileTop(btnMenuExtendet(ICNT).Tag)
                ElseIf GetFileTop(btnMenuExtendet(ICNT).Tag) + iTop < 0 Then
                    iTop = -GetFileTop(btnMenuExtendet(ICNT).Tag)
                End If
            ElseIf ICNT = 2 Then
                For IcNT2 = 1 To ListCount
                    iFile(IcNT2).ShowDetails = True
                Next IcNT2
                gen = True
            ElseIf ICNT = 3 Then
                For IcNT2 = 1 To ListCount
                    iFile(IcNT2).ShowDetails = False
                Next IcNT2
                gen = True
            End If
        End If
    Next ICNT
End If

ReleaseCapture
PicMnuExtendet.Visible = False
picShdw.Visible = False
btn.Visible = False
If gen = True Then UserControl_Resize

End Sub

Private Sub PicMnuExtendet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Integer

For ICNT = 1 To btnMenuExtendet.Count - 1
    If X >= 0 And X < PicMnuExtendet.Width / Screen.TwipsPerPixelX - 4 And Y >= btnMenuExtendet(ICNT).Top And Y < btnMenuExtendet(ICNT).Top + btnMenuExtendet(ICNT).Height Then
        btnMenuExtendet(ICNT).SetState "H"
    Else
        btnMenuExtendet(ICNT).SetState "N"
    End If
Next ICNT

End Sub


Private Sub PicMnuExtendet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PicMnuExtendet.Visible = True Then
    Dim Rec As RECT
    GetWindowRect picOzadje.hwnd, Rec
    Dim ix As Integer
    Dim iy As Integer
    
    ix = PicMnuExtendet.Left / Screen.TwipsPerPixelX - Rec.Left + X
    iy = PicMnuExtendet.Top / Screen.TwipsPerPixelY - Rec.Top + Y
    
    If ix < 0 Or iy < 0 Or iy > picOzadje.Height Or ix > picOzadje.Width Then
        PicMnuExtendet_MouseDown 0, 0, -1, -1
    End If
End If

End Sub

Private Sub picOzadje_Click()
If PicMnu.Visible = True Then
   SetCapture PicMnu.hwnd
   PicMnu.Top = my
   PicMnu.Left = mx
   PicMnu.ZOrder
End If

End Sub

Private Sub picOzadje_DblClick()
On Error Resume Next
If iFile(bSelected).Extension = "#FOLDER" Then
    RaiseEvent DirSelect(iFile(bSelected).FullName)
    Path = iFile(bSelected).FullName
Else
    If bSelected > 0 Then RaiseEvent FileSelect(GetSelectedFiles)
End If

Dim isSelcted As Boolean
Dim ICNT As Integer
isSelcted = False
For ICNT = 1 To ListCount
    If iFile(ICNT).Selected = True And Len(iFile(ICNT).FullName) > 3 And iFile(ICNT).FullName <> "#MYCOMPUTER" Then
        isSelcted = True
    End If
Next ICNT
RaiseEvent DeletableItemSelected(isSelcted)

End Sub

Private Sub picOzadje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iSelected As Long
Dim ICNT As Long
Dim IcNT2 As Long
Dim Rec As RECT

If ListCount < 1 And Button = vbLeftButton Then Exit Sub
On Error Resume Next
'first we have to calculate wich item the user intented to select
If dView = 0 Then
    For ICNT = 1 To ListCount
        If GetFileTop(ICNT) <= Y - iTop And GetFileTop(ICNT + 1) > Y - iTop Then
            iSelected = ICNT
            
            If X < picDetails.Width + 3 Then
                If Button = vbLeftButton Then
                    iFile(iSelected).ShowDetails = Not iFile(iSelected).ShowDetails
                    
                    If iFile(iSelected).ShowDetails = True And GetFileTop(ICNT) > picOzadje.Height - selDetHeight - iTop Then
                        iTop = picOzadje.Height - selDetHeight - GetFileTop(ICNT)
                    ElseIf iFile(iSelected).ShowDetails = True And GetFileTop(ICNT) + iTop < 0 Then
                        iTop = -GetFileTop(ICNT)
                    End If
                    
                    If iFile(iSelected).Selected = True Then
                        UserControl_Resize
                    End If
                End If
                
            End If
            
            Exit For
        End If
    Next ICNT
ElseIf dView = 1 Or dView = 2 Then
    For ICNT = 0 To Int(picOzadje.Height) / SelHeight - 1
        For IcNT2 = 0 To ColumnCount - 1
            If IcNT2 * ColumnWidth <= X - iLeft And (IcNT2 + 1) * ColumnWidth > X - iLeft Then
                If ICNT * SelHeight <= Y And (ICNT + 1) * SelHeight > Y Then
                    iSelected = ICNT + 1 + Int((picOzadje.Height) / SelHeight) * IcNT2
                    Exit For
                End If
            End If
        Next IcNT2
        If iSelected > 0 Then Exit For
    Next ICNT
ElseIf dView = 3 Then
    For ICNT = 0 To Int(picOzadje.Width) / 96 - 1
        For IcNT2 = 0 To ColumnCount - 1
            If IcNT2 * 65 <= Y - iTop And (IcNT2 + 1) * 65 > Y - iTop Then
                If ICNT * 96 <= X And (ICNT + 1) * 96 > X Then
                    iSelected = ICNT + 1 + Int((picOzadje.Width) / 96) * IcNT2
                    Exit For
                End If
            End If
        Next IcNT2
        If iSelected > 0 Then Exit For
    Next ICNT
End If

'now we have to select items accordingly to type of selection

If dView = 0 Then
    If iSelected < 1 Then
        If GetFileTop(ListCount) <= Y - iTop And CheckHeight > Y - iTop Then
            iSelected = ListCount
        End If
    End If
End If

If Shift = 0 Or MultiSelect = False Then  'Normal type - single item selection
    If bSelected = iSelected Then GoTo ending
    For ICNT = 1 To ListCount
        If ICNT = iSelected Then
            iFile(ICNT).Selected = True
        Else
            iFile(ICNT).Selected = False
        End If
    Next ICNT
    bSelected = iSelected
    sSelected = iSelected
ElseIf Shift = 1 Then   'Shift selection
    If bSelected < 1 Then
        For ICNT = 1 To ListCount
            If ICNT = iSelected Then
                iFile(ICNT).Selected = True
            Else
                iFile(ICNT).Selected = False
            End If
        Next ICNT
        bSelected = iSelected
    Else
        If bSelected > iSelected Then
            For ICNT = 1 To ListCount
                If ICNT >= iSelected And ICNT <= bSelected Then
                    iFile(ICNT).Selected = True
                Else
                    iFile(ICNT).Selected = False
                End If
            Next ICNT
        Else
            For ICNT = 1 To ListCount
                If ICNT <= iSelected And ICNT >= bSelected Then
                    iFile(ICNT).Selected = True
                Else
                    iFile(ICNT).Selected = False
                End If
            Next ICNT
        End If
    End If

ElseIf Shift = 2 Then ' Ctrl selection
    iFile(iSelected).Selected = Not iFile(iSelected).Selected
    bSelected = iSelected
End If

List1.Clear

For ICNT = 1 To ListCount
    If iFile(ICNT).Selected = True And iFile(ICNT).Extension <> "#FOLDER" Then List1.AddItem iFile(ICNT).name
Next ICNT

RaiseEvent FilePreSelect(List1)

UserControl_Resize 'draw selections
ending:

If GetFileTop(ICNT) <= Y - iTop And GetFileTop(ICNT + 1) > Y - iTop Then
End If

If dView > 0 Or X > picDetails.Width + 2 Then
    If Button = vbRightButton Then
    
        If iFile(iSelected).ShowDetails = True And GetFileTop(ICNT) > picOzadje.Height - selDetHeight - iTop Then
            iTop = picOzadje.Height - selDetHeight - GetFileTop(ICNT)
            UserControl_Resize
        End If

        GetWindowRect UserControl.hwnd, Rec
        
        mx = (Rec.Left + X + 1) * Screen.TwipsPerPixelX
        my = (Rec.Top + Y + 1) * Screen.TwipsPerPixelY
        
        PicMnu.Top = my
        PicMnu.Left = mx
        SetUpMnu iSelected
        PicMnu.Top = -PicMnu.Height
        
        PicMnu.Visible = True
    End If
End If
Dim isSelcted As Boolean
isSelcted = False
For ICNT = 1 To ListCount
    If iFile(ICNT).Selected = True And Len(iFile(ICNT).FullName) > 3 And iFile(ICNT).FullName <> "#MYCOMPUTER" Then
        isSelcted = True
    End If
Next ICNT
RaiseEvent DeletableItemSelected(isSelcted)

End Sub

Private Sub picOzadje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Long
Dim i As Long
btn.Visible = False
picOzadje.Refresh
If dView = 0 And Button = 0 Then
    If X < picDetails.Width + 1 And X > 1 Then
        i = 0
        For ICNT = 1 To ListCount
            If GetFileTop(ICNT) <= Y - iTop And GetFileTop(ICNT + 1) > Y - iTop Then
                i = ICNT
                Exit For
            End If
        Next ICNT
        Dim sh As Integer
        If iFile(i).ShowDetails = False Then
            sh = SelHeight - 4
        Else
            sh = selDetHeight - 4
        End If
        
        If GetFileTop(i) + 2 + iTop <= Y And GetFileTop(i) + 2 + iTop + sh >= Y Then
        
        Else
            Exit Sub
        End If
        
        If i > 0 Then
            picBuff3.ForeColor = vbHighlight
            picBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
            picBuff3.Cls
            picBuff3.Picture = LoadPicture("")
            If iFile(i).ShowDetails = False Then
                picBuff3.Height = SelHeight - 4
                picBuff3.Width = picDetails.Width - 2
                
                picDetails.BackColor = picBuff3.BackColor
                BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                
                picBuff3.Line (0, 0)-(picBuff3.Width, 0)
                picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
                picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
                picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)

                Set btn.NormalImage = picBuff3.Image
                Set btn.FocusedImage = picBuff3.Image
                
                picBuff3.ForeColor = vbHighlight
                picBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

                picDetails.BackColor = picBuff3.BackColor
                BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                
                picBuff3.Line (0, 0)-(picBuff3.Width, 0)
                picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
                picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
                picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)

                Set btn.PressedImage = picBuff3.Image
                
           Else
                picBuff3.Cls
                picBuff3.Height = selDetHeight - 4
                picBuff3.Width = picDetails.Width - 2
                
                picHideDetails.BackColor = picBuff3.BackColor
                BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                
                picBuff3.Line (0, 0)-(picBuff3.Width, 0)
                picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
                picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
                picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)
                
                Set btn.NormalImage = picBuff3.Image
                Set btn.FocusedImage = picBuff3.Image
                
                picBuff3.ForeColor = vbHighlight
                picBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

                picHideDetails.BackColor = picBuff3.BackColor
                BitBlt picBuff3.hDC, (picBuff3.Width - picDetails.Width) / 2, Int((picBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picHideDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                
                picBuff3.Line (0, 0)-(picBuff3.Width, 0)
                picBuff3.Line (0, picBuff3.Height - 1)-(picBuff3.Width, picBuff3.Height - 1)
                picBuff3.Line (0, 0)-(0, picBuff3.Height - 1)
                picBuff3.Line (picBuff3.Width - 1, 0)-(picBuff3.Width - 1, picBuff3.Height - 1)

                Set btn.PressedImage = picBuff3.Image
            End If
                
                
                btn.Top = GetFileTop(i) + 2 + iTop
                btn.Left = 2
                btn.Tag = i
                btn.Visible = True
                btn.SetFocus
        
        End If
    Else
        'PicBtnMnu.visible = False
    End If
End If
If PicRename.Visible = False Then picOzadje.SetFocus

End Sub


Private Sub picOzadje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PicMnuExtendet.Visible = True Then
    If X < 0 Or Y < 0 Or X > picOzadje.Width Or Y > picOzadje.Height Then
        PicMnuExtendet_MouseDown 0, 0, -1, -1
    End If
End If

If PicMnu.Visible = True Then
    If X < 0 Or Y < 0 Or X > picOzadje.Width Or Y > picOzadje.Height Then
        PicMnu_MouseDown 0, 0, -1, -1
    End If
End If

End Sub

Private Sub picScr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Y >= picScroller.Top And Y < picScroller.Top + picScroller.Height Then
        MouseY = -picScroller.Top + Y
    Else
        MouseY = picScroller.Height / 2
        picScr_MouseMove Button, Shift, X, Y
    End If
End If

End Sub


Private Sub picScr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Y - MouseY <= btnUP.Height Then
        picScroller.Top = btnUP.Height
        picScr.Refresh
        
        If Not iTop = 0 Then
            iTop = 0
            Generate
        End If
        
    ElseIf Y - MouseY >= picOzadje.Height - btnDOWN.Height - picScroller.Height Then
        picScroller.Top = picOzadje.Height - btnDOWN.Height - picScroller.Height
        picScr.Refresh
        
        If Not iTop = picOzadje.Height - CheckHeight Then
            iTop = picOzadje.Height - CheckHeight
            Generate
        End If
        
    Else
        picScroller.Top = Y - MouseY
        picScr.Refresh
        
        If Not iTop = Int((picScroller.Top - btnUP.Height) * (picOzadje.Height - CheckHeight) / ((picOzadje.Height - btnUP.Height - btnDOWN.Height) - picScroller.Height)) Then
            iTop = Int((picScroller.Top - btnUP.Height) * (picOzadje.Height - CheckHeight) / ((picOzadje.Height - btnUP.Height - btnDOWN.Height) - picScroller.Height))
            Generate
        End If
        
    End If
End If
End Sub


Private Sub picScrH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If X >= picScrollerH.Left And X < picScrollerH.Left + picScrollerH.Width Then
        MouseX = -picScrollerH.Left + X
    Else
        MouseX = picScrollerH.Width / 2
        picScrH_MouseMove Button, Shift, X, Y
    End If
End If

End Sub

Private Sub picScrH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If X - MouseX <= btnLEFT.Width Then
        picScrollerH.Left = btnLEFT.Width
        picScrH.Refresh
        
        If Not iLeft = 0 Then
            iLeft = 0
            Generate
        End If
    ElseIf X - MouseX >= picOzadje.Width - btnRIGHT.Width - picScrollerH.Width Then
        picScrollerH.Left = picOzadje.Width - btnRIGHT.Width - picScrollerH.Width
        picScrH.Refresh
        
        If Not iLeft = picOzadje.Width - ColumnCount * picBuff.Width Then
            iLeft = picOzadje.Width - ColumnCount * picBuff.Width
            Generate
        End If
    Else
        picScrollerH.Left = X - MouseX
        picScrH.Refresh

        If Not iLeft = Int((picScrollerH.Left - btnLEFT.Width) / ((picOzadje.Width - btnLEFT.Width - btnRIGHT.Width - picScrollerH.Width) / (picOzadje.Width - (ColumnCount * ColumnWidth)))) Then
            iLeft = Int((picScrollerH.Left - btnLEFT.Width) / ((picOzadje.Width - btnLEFT.Width - btnRIGHT.Width - picScrollerH.Width) / (picOzadje.Width - (ColumnCount * ColumnWidth))))
            Generate
        End If
    End If
End If

End Sub


Private Sub picScroller_Resize()
LN.Y1 = picScroller.Height - 1
LN.Y2 = picScroller.Height - 1

End Sub


Private Sub picScrollerH_Resize()
lnH.X1 = picScrollerH.Width - 1
lnH.X2 = picScrollerH.Width - 1

End Sub


Public Sub MouseScroll(Direction As Long)
Dim cc As Integer
Dim gTop As Long
Dim K As String

If picScrH.Visible = True Then
    If Direction = -1 Then K = "L" Else K = "R"
    cc = 100
Else
    If Direction = 1 Then K = "D" Else K = "U"
    cc = SelHeight
End If
If K = "L" Or K = "R" Then GoTo LeFTrIGHT

If K = "D" Then
    If iTop > picOzadje.Height - CheckHeight - cc Then
        gTop = iTop - cc
    Else
        gTop = picOzadje.Height + CheckHeight
    End If
ElseIf K = "U" Then
    If iTop < -cc Then
        gTop = iTop + cc
    Else
        gTop = 0
    End If
End If

If gTop <> iTop Then
    iTop = gTop
    UserControl_Resize
End If
    
Exit Sub

LeFTrIGHT:
cc = 20
Dim gLeft As Long
If K = "L" Then
    If iLeft < -cc Then
        gLeft = iLeft + cc
    Else
        gLeft = 0
    End If
ElseIf K = "R" Then
    If iLeft > picOzadje.Width - ColumnCount * picBuff.Width - cc Then
        gLeft = iLeft - cc
    Else
        gLeft = picOzadje.Width - ColumnCount * picBuff.Width
    End If
End If

If iLeft <> gLeft Then
    iLeft = gLeft
    UserControl_Resize
End If

End Sub

Private Sub tmrScr_Timer()
Dim cc As Integer
Dim gTop As Long
tmrScr.Interval = 25

If tmrScr.Tag = "L" Or tmrScr.Tag = "R" Then GoTo LeFTrIGHT
cc = 15
If tmrScr.Tag = "D" Then
    If iTop > picOzadje.Height - CheckHeight - cc Then
        gTop = iTop - cc
    Else
        gTop = picOzadje.Height + CheckHeight
    End If
ElseIf tmrScr.Tag = "U" Then
    If iTop < -cc Then
        gTop = iTop + cc
    Else
        gTop = 0
    End If
End If

If gTop <> iTop Then
    iTop = gTop
    UserControl_Resize
End If
    
Exit Sub

LeFTrIGHT:
cc = 20
Dim gLeft As Long
If tmrScr.Tag = "L" Then
    If iLeft < -cc Then
        gLeft = iLeft + cc
    Else
        gLeft = 0
    End If
ElseIf tmrScr.Tag = "R" Then
    If iLeft > picOzadje.Width - ColumnCount * picBuff.Width - cc Then
        gLeft = iLeft - cc
    Else
        gLeft = picOzadje.Width - ColumnCount * picBuff.Width
    End If
End If

If iLeft <> gLeft Then
    iLeft = gLeft
    UserControl_Resize
End If

End Sub

Private Sub txtRename_LostFocus()
PicRename.Visible = False
If txtRename.Text <> iFile(PicMnu.Tag).name Then
    If txtRename.Text <> "" Then
        On Error GoTo handle
        Dim KpATH As String
        If Right(Path, 1) = "\" Then
            KpATH = Path
        Else
            KpATH = Path & "\"
        End If
        
        Name iFile(PicMnu.Tag).FullName As KpATH & txtRename.Text
        Refresh
    End If
End If

Exit Sub

handle:
'vbError = 10 - YOU CAN USE THIS FOR ERROR HANDELING

End Sub

Private Sub UserControl_Initialize()
MyMusicPath = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Music")
MyVideosPath = GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Video")

Hook picOzadje.hwnd

Set FSO = New FileSystemObject

DoNotGenerate = True
MNU_SELECT = "Select"
MNU_VIEW = "View"
MNU_COPY = "Copy"
MNU_CUT = "Cut"
MNU_PASTE = "Paste"
MNU_RENAME = "Rename"
MNU_NEWFOLDER = "New Folder..."
MNU_PROPERTYS = "Propertys"
MNU_REFRESH = "Refresh"
MNU_DELETE = "Delete"

EXTMNU_EXPANDALL = "Show All"
EXTMNU_UNEXPANDALL = "Hide All"
EXTMNU_EXPAND = "Show Details"
EXTMNU_UNEXPAND = "Hide Details"

CAPTION_LASTACCESSED = "Last used: "
CAPTION_FILESIZE = "File Size: "
CAPTION_FILETYPE = "File Type: "
CAPTION_FOLDERPATH = "Full Name: "
CAPTION_FOLDERSIZE = "Folder Size: "
CAPTION_FILESINFOLDER = "Number of Files: "
CAPTION_DRIVEFREESPACE = "Free Space: "
CAPTION_DRIVESIZE = "Size: "
CAPTION_FILESYSTEM = "File System: "
CAPTION_DEVICEUNAVAILABLE = "Device Unavailable"

MNUVIEW_ETENDETMODE = "Extended"
MNUVIEW_LIST_SMALL = "List (Small Icons)"
MNUVIEW_LIST_LARGE = "List (Large Icons)"
MNUVIEW_ICONS = "Icons"

SelBorderColor = vbHighlight
SelBackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
SelForeColor = vbBlack
BackColor = vbWhite
ForeColor = vbBlack
FontSize = 10

picOzadje.Top = 1
picOzadje.Left = 1

picScr.Top = 1
picScrH.Left = 1

btnUP.Left = 0
btnUP.Top = 0

DrawScrollers
UserControl.BackColor = AlphaBlend(AlphaBlend(vbButtonFace, vbWhite, 170), vb3DDKShadow, 70)
btnDOWN.Left = 0
DoNotGenerate = True ' to avoid calculating when changing path
Path = "C:\"
'Generate

'we set pictures that we will use as menus
SetWindowLong PicMnuExtendet.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
SetParent PicMnuExtendet.hwnd, 0

SetWindowLong PicMnu.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
SetParent PicMnu.hwnd, 0

SetWindowLong picShdw.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
SetParent picShdw.hwnd, 0

'SetWindowLong PicMnuViews.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
'SetParent PicMnuViews.hwnd, 0
If dView = 0 Then
    SelHeight = 18 'NormalviewHeight
    selDetHeight = 64               '"Extendet" dView Mode
ElseIf dView = 1 Then
    SelHeight = 18 'NormalviewHeight
    selDetHeight = 0                'NoExtendetMode here
ElseIf dView = 2 Then
    SelHeight = 34 'NormalviewHeight
    selDetHeight = 0                'NoExtendetMode here
End If

picScr.BackColor = AlphaBlend(vb3DHighlight, vbButtonFace, 128)
picScrH.BackColor = AlphaBlend(vb3DHighlight, vbButtonFace, 128)

End Sub

Private Sub DrawBackgrounds()
'DrawsTheBackground

    With picBackground
        .Cls
        .Width = picOzadje.Width
        If dView = 3 Then
            .Height = picBuff.Height * 2
        Else
            .Height = SelHeight * 2
            picBuff.Height = .Height / 2
        End If
        
        If dView = 1 Or dView = 2 Then

            .Width = ColumnWidth
        End If
        If dView = 0 Then picBuff.Width = .Width
    End With
    
    With picBuff
        .Cls
        .BackColor = BackColor
        .ForeColor = ForeColor
    End With
    
    BitBlt picBackground.hDC, 0, 0, picBackground.Width, picBackground.Height / 2, picBuff.hDC, 0, 0, SRCCOPY
     
    With picBuff
        .Cls
        .BackColor = SelBackColor
        .ForeColor = SelBorderColor
    End With
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
        
    BitBlt picBackground.hDC, 0, picBackground.Height / 2, picBackground.Width, picBackground.Height / 2, picBuff.hDC, 0, 0, SRCCOPY

    picBackground.Refresh
    
    If dView = 0 Then DrawExtendet
     
End Sub

Private Sub DrawExtendet()
'DrawsExtendetBackground

    With picExtendet
        .Cls
        .Width = picOzadje.Width
        .Height = selDetHeight * 2
        picBuff.Height = selDetHeight
        picBuff.Width = .Width
    End With
    
    With picBuff
        .Cls
        .BackColor = AlphaBlend(BackColor, vbBlack, 245)
        .ForeColor = AlphaBlend(BackColor, vbBlack, 225)
    End With
        
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    
    BitBlt picExtendet.hDC, 0, 0, picExtendet.Width, picExtendet.Height / 2, picBuff.hDC, 0, 0, SRCCOPY
    
    picBuff.BackColor = BackColor
    BitBlt picExtendet.hDC, 1, 1, 54, picExtendet.Height / 2 - 2, picBuff.hDC, 0, 0, SRCCOPY
    
    With picBuff
        .Cls
        .BackColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
        .ForeColor = SelBorderColor
    End With
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
        
    BitBlt picExtendet.hDC, 0, picExtendet.Height / 2, picExtendet.Width, picExtendet.Height / 2, picBuff.hDC, 0, 0, SRCCOPY
    
    picBuff.BackColor = SelBackColor
    BitBlt picExtendet.hDC, 1, picExtendet.Height / 2 + 1, 55, picExtendet.Height / 2 - 2, picBuff.hDC, 0, 1, SRCCOPY

    picExtendet.Refresh
    picBuff.Height = SelHeight
     
End Sub
Public Function CheckHeight() As Long
On Error Resume Next
'here we get the height of thelist
Dim ICNT As Integer
CheckHeight = 0
If dView = 3 Then
    CheckHeight = picBuff.Height * ColumnCount
Else
    For ICNT = 1 To ListCount
        If iFile(ICNT).ShowDetails = True Then
            CheckHeight = CheckHeight + selDetHeight
        Else
            CheckHeight = CheckHeight + SelHeight
        End If
    Next ICNT
End If

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iSelected As Long
Dim ICNT As Long

If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight And ListCount > 0 Then
    If dView = 0 Then
        If KeyCode = vbKeyDown Then
            If bSelected < ListCount Then
                iSelected = bSelected + 1
            Else
                iSelected = ListCount
            End If
        ElseIf KeyCode = vbKeyUp Then
            If bSelected > 1 Then
                iSelected = bSelected - 1
            Else
                iSelected = 1
            End If
        End If
        
        If GetFileTop(iSelected) < -iTop Then
            iTop = -GetFileTop(iSelected)
            SetScroller
        ElseIf GetFileTop(iSelected + 1) > -iTop + picOzadje.Height Then
            iTop = picOzadje.Height - GetFileTop(iSelected + 1)
            SetScroller
        End If
        If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Exit Sub
    ElseIf dView = 1 Or dView = 2 Then
        If KeyCode = vbKeyDown Then
            If bSelected < ListCount Then
                iSelected = bSelected + 1
            Else
                iSelected = ListCount
            End If
        ElseIf KeyCode = vbKeyUp Then
            If bSelected > 1 Then
                iSelected = bSelected - 1
            Else
                iSelected = 1
            End If
        ElseIf KeyCode = vbKeyLeft Then
            iSelected = bSelected
            
            If bSelected > RowsInColumnCount Then
                iSelected = bSelected - RowsInColumnCount
            End If
        ElseIf KeyCode = vbKeyRight Then
            iSelected = bSelected
            
            If bSelected <= ListCount - RowsInColumnCount Then
                iSelected = bSelected + RowsInColumnCount
            End If
        End If
        
        Dim iL As Long
        iL = Int((iSelected - 1) / RowsInColumnCount) * ColumnWidth
        
        If iL < -iLeft Then
            iLeft = -iL
            SetScrollerH
        ElseIf iL > -iLeft + picOzadje.Width - ColumnWidth Then
            iLeft = picOzadje.Width - ColumnWidth - iL
            If iL < -iLeft Then
                iLeft = -iL
            End If
            
            SetScrollerH
        End If
    ElseIf dView = 3 Then
        If KeyCode = vbKeyRight Then
            If bSelected < ListCount Then
                iSelected = bSelected + 1
            Else
                iSelected = ListCount
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If bSelected > 1 Then
                iSelected = bSelected - 1
            Else
                iSelected = 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            iSelected = bSelected
            
            If bSelected > RowsInColumnCount Then
                iSelected = bSelected - RowsInColumnCount
            End If
        ElseIf KeyCode = vbKeyDown Then
            iSelected = bSelected
            
            If bSelected <= ListCount - RowsInColumnCount Then
                iSelected = bSelected + RowsInColumnCount
            End If
        End If
                iL = Int((iSelected - 1) / RowsInColumnCount) * ColumnWidth
        frmDialog.caption = iL
        If iL < -iTop Then
            iTop = -iL
            SetScroller
        ElseIf iL > -iTop + picOzadje.Height - ColumnWidth Then
            iTop = picOzadje.Height - ColumnWidth - iL
            If iL < -iTop Then
                iTop = -iL
            End If
            
            SetScroller
        End If
        
    End If

    If Shift = 0 Or MultiSelect = False Then  'Normal type - single item selection
        If bSelected = iSelected Then Exit Sub
        For ICNT = 1 To ListCount
            If ICNT = iSelected Then
                iFile(ICNT).Selected = True
            Else
                iFile(ICNT).Selected = False
            End If
        Next ICNT
        bSelected = iSelected
        If KeyCode = vbKeyUp Or vbKeyDown Then sSelected = iSelected
            
    ElseIf Shift = 1 Then   'Shift selection
        If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then If dView = 0 Then Exit Sub
        If sSelected < 1 Then
            For ICNT = 1 To ListCount
                If ICNT = iSelected Then
                    iFile(ICNT).Selected = True
                Else
                    iFile(ICNT).Selected = False
                End If
            Next ICNT
        Else
            If sSelected > iSelected Then
                For ICNT = 1 To ListCount
                    If ICNT >= iSelected And ICNT <= sSelected Then
                        iFile(ICNT).Selected = True
                    Else
                        iFile(ICNT).Selected = False
                    End If
                Next ICNT
            Else
                For ICNT = 1 To ListCount
                    If ICNT <= iSelected And ICNT >= sSelected Then
                        iFile(ICNT).Selected = True
                    Else
                        iFile(ICNT).Selected = False
                    End If
                Next ICNT
            End If
        End If
        bSelected = iSelected
    End If
    Generate 'draw selections
ElseIf KeyCode = 13 Then
    If PicRename.Visible = True Then
        picOzadje.SetFocus
    Else
        picOzadje_DblClick
    End If
ElseIf KeyCode = vbKeyA And Shift = 2 And Me.MultiSelect = True Then
    For ICNT = 1 To ListCount
        iFile(ICNT).Selected = True
        If iFile(ICNT).Extension <> "#FOLDER" Then bSelected = ICNT
    Next ICNT
    Generate
ElseIf KeyCode = vbKeyDelete Then
    Dim ref As Boolean
    For ICNT = 1 To ListCount
        If iFile(ICNT).Selected = True Then
            If RemoveFile(iFile(ICNT).FullName, RecycleFile) = True Then
                ref = True
            End If
        End If
    Next ICNT
    If ref = True Then Refresh
ElseIf KeyCode = vbKeyF2 Then

Else
    If PicRename.Visible = False Then RaiseEvent KeyDown(KeyCode, Shift)
    
End If


End Sub

Private Sub UserControl_Resize()

DoNotGenerate = True
If dView = 0 Or dView = 3 Then
    picOzadje.Height = UserControl.Height / Screen.TwipsPerPixelY - 2
    picScrH.Visible = False
    picScr.Left = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 1
    picScr.Height = picOzadje.Height
    
    'On Error Resume Next
    If CheckHeight > UserControl.Height / Screen.TwipsPerPixelY - 2 Then
        If iTop + CheckHeight < picOzadje.Height Then iTop = picOzadje.Height - CheckHeight
        SetScroller
        btnDOWN.Top = picScr.Height - btnDOWN.Height
        picScr.Visible = True
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 2
    Else
        iTop = 0
        picScr.Visible = False
        picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    End If
ElseIf dView = 1 Or dView = 2 Then
    picOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    
    picScr.Visible = False
    picScrH.Top = UserControl.Height / Screen.TwipsPerPixelY - picScrH.Height - 1
    picScrH.Width = picOzadje.Width
End If
DoNotGenerate = False
Generate

End Sub

Private Sub SetScroller()
On Error Resume Next
If picOzadje.Height * (picOzadje.Height - btnUP.Height - btnDOWN.Height) / (CheckHeight) > 10 Then
    picScroller.Height = picOzadje.Height * (picOzadje.Height - btnUP.Height - btnDOWN.Height) / (CheckHeight)
Else
    picScroller.Height = 10
End If

If Not picScroller.Top = (iTop) * (picOzadje.Height - btnUP.Height - btnDOWN.Height - picScroller.Height) / (picOzadje.Height - (CheckHeight)) + btnUP.Height Then picScroller.Top = (iTop) * (picOzadje.Height - btnUP.Height - btnDOWN.Height - picScroller.Height) / (picOzadje.Height - (CheckHeight)) + btnUP.Height

picScr.Refresh



End Sub

Private Sub SetScrollerH()
On Error Resume Next
If picOzadje.Width * (picOzadje.Width - btnLEFT.Width - btnRIGHT.Width) / (ColumnCount * picBuff.Width) > 10 Then
    picScrollerH.Width = picOzadje.Width * (picOzadje.Width - btnLEFT.Width - btnRIGHT.Width) / (ColumnCount * picBuff.Width)
Else
    picScrollerH.Width = 10
End If

picScrollerH.Left = (iLeft) * (picOzadje.Width - btnLEFT.Width - btnRIGHT.Width - picScrollerH.Width) / (picOzadje.Width - (ColumnCount * picBuff.Width)) + btnLEFT.Width

picScrH.Refresh
End Sub

Public Property Get Path() As String
    Path = dPath
End Property

Public Function GetSelectedFiles() As ListBox
'returns all the selected files to the user
Dim ICNT As Integer
lstFiles.Clear

For ICNT = 1 To ListCount
    If iFile(ICNT).Selected = True And iFile(ICNT).Extension <> "#FOLDER" Then
        lstFiles.AddItem iFile(ICNT).FullName
    End If
Next ICNT

Set GetSelectedFiles = lstFiles

End Function

Public Sub Refresh()
qk = True
Dim cPath As String
cPath = dPath
Dir1.Refresh
dPath = ""
Path = cPath
DrawBackgrounds
qk = False

End Sub

Public Property Let View(NewView As Integer)
iLeft = 0
iTop = 0
dView = NewView
qk = True
If DoNotGenerate = False Then UserControl_Resize 'Generate '
'If DoNotGenerate = False Then Me.Refresh
qk = False
RaiseEvent ViewChange(NewView)

End Property

Public Property Get View() As Integer
View = dView

End Property

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

Private Sub UserControl_Terminate()
UnHook picOzadje.hwnd

End Sub


