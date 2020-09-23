VERSION 5.00
Begin VB.Form frmMKDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMKDir"
   ClientHeight    =   1035
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5805
   ClipControls    =   0   'False
   Icon            =   "frmMKDir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   330
      Left            =   4440
      TabIndex        =   2
      Top             =   555
      Width           =   1215
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
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picFile 
      BackColor       =   &H00996422&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   120
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
      Width           =   4155
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
         Left            =   720
         TabIndex        =   0
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox picFileBack 
         BackColor       =   &H00996422&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   960
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   3555
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "lbl"
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
      TabIndex        =   5
      Top             =   240
      Width           =   150
   End
End
Attribute VB_Name = "frmMKDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()
On Error GoTo handle
If txtFile.Text <> "" Then
    If Right(frmDialog.FileView.Path, 1) = "\" Then
        MkDir frmDialog.FileView.Path & txtFile.Text
    Else
        MkDir frmDialog.FileView.Path & "\" & txtFile.Text
    End If

    frmDialog.FileView.Refresh
Else
    MsgBox "You must type a name!", vbExclamation
    Exit Sub
End If
Unload Me
Exit Sub
handle:

If vbError = 10 Then
    MsgBox "Illegal name, or folder with this name already exists!" & vbNewLine & "Try with another name.", vbExclamation
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
picFile.BackColor = frmDialog.picFile.BackColor
picFileBack.Top = 1
picFileBack.Left = 1
picFileBack.Width = picFile.Width - 2
picFileBack.Height = picFile.Height - 2
picFileBack.BackColor = Me.txtFile.BackColor

Me.txtFile.Left = 2
Me.txtFile.Width = picFile.Width - 4
Me.txtFile.Top = (Me.picFile.Height - Me.txtFile.Height) / 2 + 1

Me.caption = iCAPTION_MKDIR_TITLE
lbl.caption = iCAPTION_MKDIR_LABLE
cmdOK.caption = iCAPTION_MKDIR_OK
cmdCancel.caption = iCAPTION_MKDIR_CANCEL
txtFile.Text = iCAPTION_MKDIR_DEFAULTFOLDER

End Sub

Private Sub lblFile_Click()

End Sub

