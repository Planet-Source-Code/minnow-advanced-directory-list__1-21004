VERSION 5.00
Begin VB.UserControl UserControl1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1008
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   984
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1008
   ScaleWidth      =   984
   Begin VB.PictureBox Picture1 
      Height          =   612
      Left            =   0
      Picture         =   "Directory List.ctx":0000
      ScaleHeight     =   564
      ScaleWidth      =   564
      TabIndex        =   0
      Top             =   0
      Width           =   612
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub UserControl_Initialize()

  DirListDone = False
  
End Sub

Public Property Let ShowRecurseCheckbox(v_blnSRC As Boolean)

  DirListDone = False
  ShowRecurseFolders = v_blnSRC

End Property

Public Property Get ShowRecurseCheckbox() As Boolean

  DirListDone = False
  ShowRecurseCheckbox = ShowRecurseFolders

End Property

Public Property Let RecurseChecked(v_blnRC As Boolean)

  DirListDone = False
  RecurseFolders = v_blnRC

End Property

Public Property Get RecurseChecked() As Boolean

  DirListDone = False
  RecurseChecked = RecurseFolders

End Property

Public Function ShowDirectoyList() As Boolean

  DirListDone = False
  DirectoryList.Show

End Function

Public Property Get DirectoryListDone() As Boolean

  DirectoryListDone = DirListDone

End Property

Public Property Get SelectedPath() As String

  SelectedPath = SelectedFolder

End Property

Public Property Get Cancelled() As Boolean

  Cancelled = CancelClicked

End Property

Private Sub UserControl_Resize()

  UserControl.Width = Picture1.Width + 40
  UserControl.Height = Picture1.Height + 40
  
End Sub

