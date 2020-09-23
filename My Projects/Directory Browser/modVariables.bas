Attribute VB_Name = "modVariables"
Option Explicit

'Variables for passing data back and forth from the Form to the Class Module
Public ShowRecurseFolders As Boolean
Public RecurseFolders As Boolean
Public CancelClicked As Boolean
Public SelectedFolder As String
Public DirListDone As Boolean
'

Private Sub Main()

  ShowRecurseFolders = True
  RecurseFolders = True
  CancelClicked = False
  SelectedFolder = ""
  
  DirectoryList.Show

End Sub
