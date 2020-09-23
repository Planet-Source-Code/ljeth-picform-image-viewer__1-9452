VERSION 5.00
Begin VB.Form frmPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Path"
   ClientHeight    =   3990
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.DriveListBox drvDriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.DirListBox dirDirList 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4440
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4665
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4665
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   1215
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
  Unload Me
End Sub

Private Sub drvDriveList_Change()

  On Error GoTo endSub
  
  'change path with drive
  dirDirList.path = drvDriveList.Drive
  
endSub:

End Sub

Private Sub Form_Load()

  'get current path from main form
  dirDirList.path = frmPicForm.gPath
  
End Sub

Private Sub OKButton_Click()

  'store path in main form and exit
  frmPicForm.gPath = dirDirList.path
  frmPicForm.filImgList.path = frmPicForm.gPath
  frmPicForm.cmdPrev.Enabled = False
  frmPicForm.cmdNext.Enabled = False
  Unload Me
  
End Sub
