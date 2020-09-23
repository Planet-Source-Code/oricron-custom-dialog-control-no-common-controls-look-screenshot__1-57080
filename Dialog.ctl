VERSION 5.00
Begin VB.UserControl Dialog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   ForeColor       =   &H00996422&
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   52
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dialog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00996422&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00996422&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++
'Author: Oricron

'Project: Dialog Control

'Short Description: A common dialog style control for
'open and save files...

'Comment: The code that is not mine is commented as i
'found it if it is changed I apologize to the authors
'- contact me to give you credits etc...
'++++++++++++++++++++++++++++++++++++++++++++++++++++

'Local Variables...
Dim dTitle As String
Dim dInitDir As String
Dim CMinWidth As Integer
Dim CMinHeight As Integer

Dim cBtnCancel As String
Dim cBtnOpen As String
Dim cBtnSave As String
Dim dMultiSelect As Boolean
Dim dSizable As Boolean
Dim dLblBrowseForFile As String
Dim dLblFile As String
Dim dLblExtensions As String
Dim dRememberDirectory As Boolean

Public CAPTION_EXTENDETVIEW As String
Public CAPTION_LIST_SMALLICONS As String
Public CAPTION_LIST_LARGEICONS As String
Public CAPTION_ICONS As String

Public CAPTION_MKDIR_TITLE As String
Public CAPTION_MKDIR_LABLE As String
Public CAPTION_MKDIR_OK As String
Public CAPTION_MKDIR_CANCEL As String
Public CAPTION_MKDIR_DEFAULTFOLDER As String

Public EXTMNU_EXPANDALL As String
Public EXTMNU_UNEXPANDALL As String
Public EXTMNU_EXPAND As String
Public EXTMNU_UNEXPAND As String

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

Public MNU_SELECT As String
'Public MNU_COPY As String
'Public MNU_CUT As String
'Public MNU_PASTE As String
Public MNU_RENAME As String
Public MNU_PROPERTYS As String
Public MNU_REFRESH As String
Public MNU_DELETE As String

'eVENTS
Public Event FilesSelected(iFiles As oList)
Public Event DialogCanceled()
Public Event SaveFileAs(FileName As String)

Public FileName As String

Public Sub Init()
frmDialog.FileView.DoNotGenerate = True
Load frmDialog
End Sub

Private Sub UserControl_Initialize()
DialogWidth = (510) * Screen.TwipsPerPixelX
DialogHeight = (350) * Screen.TwipsPerPixelY
CMinWidth = DialogWidth
CMinHeight = DialogHeight
frmDialog.MinHeight = CMinHeight
frmDialog.MinWidth = CMinWidth

dDialogFilter = "All Files|*.*"

CAPTION_EXTENDETVIEW = "Extended"
CAPTION_LIST_SMALLICONS = "List (Small Icons)"
CAPTION_LIST_LARGEICONS = "List (Large Icons)"
CAPTION_ICONS = "Icons"

CAPTION_MKDIR_TITLE = "New Folder..."
CAPTION_MKDIR_LABLE = "Name New Folder as:"
CAPTION_MKDIR_OK = "&OK"
CAPTION_MKDIR_CANCEL = "&Cancel"
CAPTION_MKDIR_DEFAULTFOLDER = "New Folder"

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

MNU_SELECT = "Select"
'MNU_VIEW = "View"
'MNU_COPY = "Copy"
'MNU_CUT = "Cut"
'MNU_PASTE = "Paste"
MNU_RENAME = "Rename"
MNU_PROPERTYS = "Propertys..."
MNU_REFRESH = "Refresh"
MNU_DELETE = "Delete"

DialogView = 1
'set the width and height of the dialog (it is also the minimum)

dTitle = "Dialog Control v 1.0"

cBtnCancel = "Cancel"
cBtnOpen = "Open"
cBtnSave = "Save"
dLblBrowseForFile = "Browse for File: "
dLblFile = "File Name: "
dLblExtensions = "File Type: "

'Nothing useful, just ui for IDE
Shape1.Top = 0
Shape1.Left = 0
Shape1.Width = 44
Shape1.Height = 32

Label1.caption = "Dialog" & vbNewLine & "v 1.0"

Label1.Left = (44 - Label1.Width) / 2
Label1.Top = (32 - Label1.Height) / 2

On Error Resume Next
frmDialog.FileView.Generate

End Sub

Private Sub UserControl_Resize()
On Error Resume Next

UserControl.Width = 44 * Screen.TwipsPerPixelX
UserControl.Height = 32 * Screen.TwipsPerPixelY

End Sub

'Sets the dialog title (before Show!!!)
Public Property Let DialogTitle(NewTitle As String)
dTitle = NewTitle


End Property

Public Property Get DialogTitle() As String
DialogTitle = dTitle

End Property

'Set Startup Folder
Public Property Let InitDir(NewPath As String)
dInitDir = NewPath

End Property

Public Property Get InitDir() As String
InitDir = dInitDir

End Property

'To alow MultiSelect during dialog open
Public Property Let MultiSelect(NewValue As Boolean)
dMultiSelect = NewValue

End Property

Public Property Get MultiSelect() As Boolean
MultiSelect = dMultiSelect

End Property

'To alow resizing of dialog
Public Property Let Sizable(NewValue As Boolean)
dSizable = NewValue

End Property

Public Property Get Sizable() As Boolean
Sizable = dSizable

End Property

'If true when you second open the same directory appears else it puts you back into initial dir
Public Property Let RememberCurrentPath(Value As Boolean)
dRememberDirectory = Value

End Property

Public Property Get RememberCurrentPath() As Boolean
RememberCurrentPath = dRememberDirectory

End Property

'Sets the dialog dLblBrowseForFile CAPTION (before Show!!!)
Public Property Let Dialog_LABELBROWSE_CAPTION(NewTitle As String)
dLblBrowseForFile = NewTitle

End Property

Public Property Get Dialog_LABELBROWSE_CAPTION() As String
Dialog_LABELBROWSE_CAPTION = dLblBrowseForFile

End Property


'Sets DialogFilter(before Show!!!)
Public Property Let DialogFilter(filter As String)
dDialogFilter = filter
frmDialog.SetMnuExtensions
frmDialog.shpExtensionsSel.Tag = 1

End Property

Public Property Get DialogFilter() As String
DialogFilter = dDialogFilter

End Property

'Sets the dialog dLblFile CAPTION (before Show!!!)
Public Property Let Dialog_LABELFILE_CAPTION(NewTitle As String)
dLblFile = NewTitle

End Property

Public Property Get Dialog_LABELFILE_CAPTION() As String
Dialog_LABELFILE_CAPTION = dLblFile

End Property

'Sets the dialog dLblExtensions CAPTION (before Show!!!)
Public Property Let Dialog_LABELFILETYPE_CAPTION(NewTitle As String)
dLblExtensions = NewTitle

End Property

Public Property Get Dialog_LABELFILETYPE_CAPTION() As String
Dialog_LABELFILETYPE_CAPTION = dLblExtensions

End Property

'Sets the dialog cBtnCancel CAPTION (before Show!!!)
Public Property Let Dialog_CANCELBUTTON_CAPTION(NewTitle As String)
cBtnCancel = NewTitle

End Property

Public Property Get Dialog_CANCELBUTTON_CAPTION() As String
Dialog_CANCELBUTTON_CAPTION = cBtnCancel

End Property

'Sets the dialog cBtnOpen CAPTION (before Show!!!)
Public Property Let Dialog_OPENBUTTON_CAPTION(NewTitle As String)
cBtnOpen = NewTitle

End Property

Public Property Get Dialog_OPENBUTTON_CAPTION() As String
Dialog_OPENBUTTON_CAPTION = cBtnOpen

End Property

'Sets the dialog cBtnSave CAPTION (before Show!!!)
Public Property Let Dialog_SAVEBUTTON_CAPTION(NewTitle As String)
cBtnSave = NewTitle

End Property

Public Property Get Dialog_SAVEBUTTON_CAPTION() As String
Dialog_SAVEBUTTON_CAPTION = cBtnSave

End Property

Public Sub ShowOpen()
On Error Resume Next
frmDialog.FileView.DoNotGenerate = True
'SetUpThe Dialog
frmDialog.caption = dTitle
iCAPTION_MKDIR_TITLE = CAPTION_MKDIR_TITLE
iCAPTION_MKDIR_LABLE = CAPTION_MKDIR_LABLE
iCAPTION_MKDIR_OK = CAPTION_MKDIR_OK
iCAPTION_MKDIR_CANCEL = CAPTION_MKDIR_CANCEL
iCAPTION_MKDIR_DEFAULTFOLDER = CAPTION_MKDIR_DEFAULTFOLDER

frmDialog.FileView.EXTMNU_EXPANDALL = EXTMNU_EXPANDALL
frmDialog.FileView.EXTMNU_UNEXPANDALL = EXTMNU_UNEXPANDALL
frmDialog.FileView.EXTMNU_EXPAND = EXTMNU_EXPAND
frmDialog.FileView.EXTMNU_UNEXPAND = EXTMNU_UNEXPAND

frmDialog.FileView.CAPTION_LASTACCESSED = CAPTION_LASTACCESSED
frmDialog.FileView.CAPTION_FILESIZE = CAPTION_FILESIZE
frmDialog.FileView.CAPTION_FILETYPE = CAPTION_FILETYPE
frmDialog.FileView.CAPTION_FOLDERPATH = CAPTION_FOLDERPATH
frmDialog.FileView.CAPTION_FOLDERSIZE = CAPTION_FOLDERSIZE
frmDialog.FileView.CAPTION_FILESINFOLDER = CAPTION_FILESINFOLDER
frmDialog.FileView.CAPTION_DRIVEFREESPACE = CAPTION_DRIVEFREESPACE
frmDialog.FileView.CAPTION_DRIVESIZE = CAPTION_DRIVESIZE
frmDialog.FileView.CAPTION_FILESYSTEM = CAPTION_FILESYSTEM
frmDialog.FileView.CAPTION_DEVICEUNAVAILABLE = CAPTION_DEVICEUNAVAILABLE

frmDialog.FileView.MNUVIEW_ETENDETMODE = CAPTION_EXTENDETVIEW
frmDialog.FileView.MNUVIEW_LIST_SMALL = CAPTION_LIST_SMALLICONS
frmDialog.FileView.MNUVIEW_LIST_LARGE = CAPTION_LIST_LARGEICONS
frmDialog.FileView.MNUVIEW_ICONS = CAPTION_ICONS
frmDialog.FileView.MNU_NEWFOLDER = CAPTION_MKDIR_TITLE
frmDialog.FileView.MNU_SELECT = MNU_SELECT
frmDialog.FileView.MNU_RENAME = MNU_RENAME
frmDialog.FileView.MNU_PROPERTYS = MNU_PROPERTYS
frmDialog.FileView.MNU_REFRESH = MNU_REFRESH
frmDialog.FileView.MNU_DELETE = MNU_DELETE

If dRememberDirectory = True And DialogPath <> "" Then
    If frmDialog.FileView.Path <> DialogPath Then
        frmDialog.FileView.Path = DialogPath
    End If
Else
    If InitDir <> "" Then
        If frmDialog.FileView.Path <> dInitDir Then
            frmDialog.FileView.Path = dInitDir
        End If
    End If
End If

'Referesh shown folder
'frmDialog.FileView.Refresh

'Clear the Back Forward lists
frmDialog.lstBack.Clear
frmDialog.lstForward.Clear

frmDialog.mnuBtn(frmDialog.iBtnBackIndex).Enabled = False
'frmDialog.Width = DialogWidth
'frmDialog.Height = DialogHeight
frmDialog.MinHeight = CMinHeight
frmDialog.MinWidth = CMinWidth

frmDialog.cmdOK.caption = cBtnOpen
frmDialog.cmdCancel.caption = cBtnCancel

frmDialog.cmdOK.Height = 22
frmDialog.cmdCancel.Height = 22
frmMKDir.cmdOK.Height = 22
frmMKDir.cmdCancel.Height = 22
frmDialog.cmdCancel.Top = 27

frmDialog.picButtons.Width = frmDialog.cmdOK.Width
frmDialog.picButtons.Height = frmDialog.cmdCancel.Top + frmDialog.cmdCancel.Height

frmDialog.picBottom.Height = frmDialog.picButtons.Height + 15 + frmDialog.picButtons.Top


frmDialog.FileView.MultiSelect = dMultiSelect
frmDialog.iSizable = dSizable
frmDialog.picResize.Visible = dSizable

'Sets label captions
frmDialog.lblBrowse.caption = dLblBrowseForFile
frmDialog.lblExtensions.caption = dLblExtensions
frmDialog.lblFile.caption = dLblFile

'set menu captions
frmDialog.iView0 = CAPTION_EXTENDETVIEW
frmDialog.iView1 = CAPTION_LIST_SMALLICONS
frmDialog.iView2 = CAPTION_LIST_LARGEICONS
frmDialog.iView3 = CAPTION_ICONS

'set last view

frmDialog.FileView.View = DialogView
'frmDialog.FileView.DoNotGenerate = False
'Show the dialog

'Clear the Back Forward lists

frmDialog.FileView.DoNotGenerate = True
Dim asd() As String
asd() = Split(dDialogFilter, "|")

frmDialog.lblExtensionsCap.caption = asd(0)
frmDialog.FileView.View = DialogView
frmDialog.FileView.DoNotGenerate = False
frmDialog.FileView.filter = asd(1)

frmDialog.lstBack.Clear
frmDialog.lstForward.Clear

frmDialog.mnuBtn(frmDialog.iBtnBackIndex).Enabled = False

DoNotUnload = True

frmDialog.SetMnuExtensions
frmDialog.txtFile.Text = ""
frmDialog.FileView.Resize

frmDialog.Show vbModal
FileName = frmDialog.txtFile.Text

If dCancel = True Then
    RaiseEvent DialogCanceled
Else
    If frmDialog.lstFiles.ListCount > 0 Then
        RaiseEvent FilesSelected(frmDialog.lstFiles)
    End If
End If
DoNotUnload = False
'Unload frmDialog

End Sub

Public Sub ShowSave(Optional FileName As String)
frmDialog.FileView.DoNotGenerate = True
'SetUpThe Dialog
frmDialog.caption = dTitle

iCAPTION_MKDIR_TITLE = CAPTION_MKDIR_TITLE
iCAPTION_MKDIR_LABLE = CAPTION_MKDIR_LABLE
iCAPTION_MKDIR_OK = CAPTION_MKDIR_OK
iCAPTION_MKDIR_CANCEL = CAPTION_MKDIR_CANCEL
iCAPTION_MKDIR_DEFAULTFOLDER = CAPTION_MKDIR_DEFAULTFOLDER

frmDialog.FileView.EXTMNU_EXPANDALL = EXTMNU_EXPANDALL
frmDialog.FileView.EXTMNU_UNEXPANDALL = EXTMNU_UNEXPANDALL
frmDialog.FileView.EXTMNU_EXPAND = EXTMNU_EXPAND
frmDialog.FileView.EXTMNU_UNEXPAND = EXTMNU_UNEXPAND

frmDialog.FileView.CAPTION_LASTACCESSED = CAPTION_LASTACCESSED
frmDialog.FileView.CAPTION_FILESIZE = CAPTION_FILESIZE
frmDialog.FileView.CAPTION_FILETYPE = CAPTION_FILETYPE
frmDialog.FileView.CAPTION_FOLDERPATH = CAPTION_FOLDERPATH
frmDialog.FileView.CAPTION_FOLDERSIZE = CAPTION_FOLDERSIZE
frmDialog.FileView.CAPTION_FILESINFOLDER = CAPTION_FILESINFOLDER
frmDialog.FileView.CAPTION_DRIVEFREESPACE = CAPTION_DRIVEFREESPACE
frmDialog.FileView.CAPTION_DRIVESIZE = CAPTION_DRIVESIZE
frmDialog.FileView.CAPTION_FILESYSTEM = CAPTION_FILESYSTEM
frmDialog.FileView.CAPTION_DEVICEUNAVAILABLE = CAPTION_DEVICEUNAVAILABLE

frmDialog.FileView.MNUVIEW_ETENDETMODE = CAPTION_EXTENDETVIEW
frmDialog.FileView.MNUVIEW_LIST_SMALL = CAPTION_LIST_SMALLICONS
frmDialog.FileView.MNUVIEW_LIST_LARGE = CAPTION_LIST_LARGEICONS
frmDialog.FileView.MNUVIEW_ICONS = CAPTION_ICONS
frmDialog.FileView.MNU_NEWFOLDER = CAPTION_MKDIR_TITLE
frmDialog.FileView.MNU_SELECT = MNU_SELECT
frmDialog.FileView.MNU_RENAME = MNU_RENAME
frmDialog.FileView.MNU_PROPERTYS = MNU_PROPERTYS
frmDialog.FileView.MNU_REFRESH = MNU_REFRESH
frmDialog.FileView.MNU_DELETE = MNU_DELETE
'frmDialog.FileView.Height = frmDialog.ScaleHeight - frmDialog.FileView.Top - frmDialog.picBottom.Height
'frmDialog.FileView.Width = frmDialog.ScaleWidth - frmDialog.FileView.Left - frmDialog.picRight.Width

If dRememberDirectory = True And DialogPath <> "" Then
    If frmDialog.FileView.Path <> DialogPath Then
        frmDialog.FileView.Path = DialogPath
    End If
Else
    If InitDir <> "" Then
        If frmDialog.FileView.Path <> dInitDir Then
            frmDialog.FileView.Path = dInitDir
        End If
    End If
End If

'Referesh shown folder
'frmDialog.FileView.Refresh

frmDialog.Width = DialogWidth
frmDialog.Height = DialogHeight
frmDialog.MinHeight = CMinHeight
frmDialog.MinWidth = CMinWidth

frmDialog.cmdOK.caption = cBtnSave
frmDialog.cmdCancel.caption = cBtnCancel

frmDialog.cmdOK.Height = 22
frmDialog.cmdCancel.Height = 22
frmDialog.cmdCancel.Top = 27

frmDialog.picButtons.Width = frmDialog.cmdOK.Width
frmDialog.picButtons.Height = frmDialog.cmdCancel.Top + frmDialog.cmdCancel.Height

frmDialog.picBottom.Height = frmDialog.picButtons.Height + 15 + frmDialog.picButtons.Top

frmDialog.FileView.MultiSelect = False
frmDialog.iSizable = dSizable
frmDialog.picResize.Visible = dSizable

'Sets label captions
frmDialog.lblBrowse.caption = dLblBrowseForFile
frmDialog.lblExtensions.caption = dLblExtensions
frmDialog.lblFile.caption = dLblFile

'set menu captions
frmDialog.iView0 = "Extended"
frmDialog.iView1 = "List (Small Icons)"
frmDialog.iView2 = "List (Large Icons)"
frmDialog.iView3 = "Icons"

'set last view

frmDialog.FileView.View = DialogView
'frmDialog.FileView.DoNotGenerate = False
'Show the dialog
frmDialog.FileView.DoNotGenerate = True
Dim asd() As String
asd() = Split(dDialogFilter, "|")

frmDialog.FileView.filter = asd(1)
frmDialog.lblExtensionsCap.caption = asd(0)

frmDialog.FileView.DoNotGenerate = False

frmDialog.FileView.Refresh
frmDialog.txtFile.Text = FileName
frmDialog.SaveMode = True

'Clear the Back Forward lists

frmDialog.lstBack.Clear
frmDialog.lstForward.Clear

frmDialog.mnuBtn(frmDialog.iBtnBackIndex).Enabled = False

DoNotUnload = True 'WILL CANCEL UNLOAD OF FRMDIALOG WHEN USER PRESSES THE X BUTTON...SO THE NEXT LOAD IS MUCH FASTER

frmDialog.SetMnuExtensions

frmDialog.Show vbModal
FileName = frmDialog.txtFile.Text
If dCancel = True Then
    RaiseEvent DialogCanceled
Else
    If frmDialog.lstFiles.ListCount > 0 Then
        RaiseEvent SaveFileAs(frmDialog.lstFiles.GetItem(0))
    
    End If
End If

frmDialog.SaveMode = False

DoNotUnload = False



End Sub

Public Sub UserControl_Terminate()
On Error Resume Next
Unload frmDialog
Unload frmMKDir

End Sub

Public Sub Terminate()
DoNotUnload = False
Unload frmDialog
Unload frmMKDir

End Sub

