VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DirectoryList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Directory List"
   ClientHeight    =   4788
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   3132
   Icon            =   "Directory List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4788
   ScaleWidth      =   3132
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   4200
      Width           =   972
   End
   Begin VB.CheckBox chkRecurse 
      Caption         =   "Recurse Subfolders"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5892
   End
   Begin MSComctlLib.ImageList IconList 
      Left            =   6240
      Top             =   120
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":0ECA
            Key             =   "UnknownDrive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":1464
            Key             =   "FloppyDrive"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":19FE
            Key             =   "FixedDrive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":1F98
            Key             =   "NetworkDrive"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":2532
            Key             =   "CDROMDrive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":2ACC
            Key             =   "RAMDrive"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":3066
            Key             =   "Desktop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":3600
            Key             =   "MyComputer"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":3B9A
            Key             =   "MyDocuments"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":4134
            Key             =   "NetworkNeighbourhood"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":46CE
            Key             =   "ClosedFolder"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Directory List.frx":4C68
            Key             =   "OpenFolder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView DS 
      Height          =   3732
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   6583
      _Version        =   393217
      Indentation     =   176
      LineStyle       =   1
      Style           =   7
      ImageList       =   "IconList"
      Appearance      =   1
   End
End
Attribute VB_Name = "DirectoryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FSO As New Scripting.FileSystemObject 'Form's Global Declarations
Dim NS As New clsNodeCol
Const FormMinWidth = 2500
Const FormMinHeight = 2500
'
'******************  My Accronyms  ******************
'FSO = File System Object (File Object)
'NS = Node Structure (Class Collection Object)
'DS = Directory Structure (Treeview Object)
'

'******************  Events  ******************
'Raised on Cancel button click
Private Sub cmdCancel_Click()

  'Flag the button as being clicked
  CancelClicked = True
  
  'Variable to determine if the user is done with the directory list
  DirListDone = True
  
  'Clear the Selected Folder Variable
  SelectedFolder = ""
  
  'Unload the form
  Unload Me

End Sub

'Raised on OK button click
Private Sub cmdOK_Click()

  'Set the RecurseFolders flag
  If chkRecurse.Value = 1 Then
    RecurseFolders = True
  Else
    RecurseFolders = False
  End If

  'Variable to determine if the user is done with the directory list
  DirListDone = True
  
  'Unload the form
  Unload Me

End Sub

'Raised on treeview expansion
Private Sub DS_Expand(ByVal Node As MSComctlLib.Node)
  
  'If the folder has already been populated then there is no sense in
  'repopulating it.
  If Not NS.Item(Node.Key).HasBeenBrowsed Then
    'We create a dummy node if the current node has children, new we must
    'remove it. This is explained in greater detail later.
    RemoveChildren Node.Key
    'Populate the node.
    AddChildren Node.Key
    'Set the node's HasBeenBrowsed property so that we don't try to populate
    'it again later.
    NS.Item(Node.Key).HasBeenBrowsed = True
    
  End If
        
End Sub

Private Sub DS_NodeClick(ByVal Node As MSComctlLib.Node)

  'We need to turn off the OK button if we can't produce a valid folder.
  cmdOK.Enabled = NS.Item(Node.Key).CanClickOK
  
  'Set the selected folder to the current folder
  SelectedFolder = NS.Item(DS.SelectedItem.Key).FolderPath

End Sub

'Raise on Form Load
Private Sub Form_Load()
  
  'Setup everything
  NS.Clear
  DirListDone = False
  InitialPosition
  PopulateTree
   
End Sub

Private Sub Form_Resize()

  'If the window is minimized (which shouldn't ever be an option but if
  'the window state changes for some reason the program will crash) we
  'want to exit the function.
  If DirectoryList.WindowState = vbMinimized Then Exit Sub
  
  'Sets up the minimum values for the window size
  If DirectoryList.Width < FormMinWidth Then DirectoryList.Width = FormMinWidth
  If DirectoryList.Height < FormMinHeight Then DirectoryList.Height = FormMinHeight
  
  ResizeEverything

End Sub

'******************  Private Functions  ******************

Private Function InitialPosition()

  Dim DSTop As Integer
  
  'Set the DS Top to 120 by default
  DSTop = 120
  
  'Turn the Recurse checkbox visible on/off
  chkRecurse.Visible = ShowRecurseFolders
  'Sets the default option
  If RecurseFolders Then chkRecurse.Value = 1
  
  'If we are showing the checkbox, we need to make room for it.
  If ShowRecurseFolders Then DSTop = 370
  
  'Set the initial positions
  DS.Left = 120
  DS.Top = DSTop
  chkRecurse.Left = 120
  chkRecurse.Top = 0
  DirectoryList.Width = 4000
  DirectoryList.Height = 6000
  
  ResizeEverything
    
End Function

Private Function ResizeEverything()

  'Sets the widths, heights, and positions where needed
  DS.Width = DirectoryList.Width - 350
  DS.Height = DirectoryList.Height - (DS.Top + 1000)
  cmdCancel.Left = (120 + DS.Width) - cmdCancel.Width
  cmdCancel.Top = DS.Top + DS.Height + 120
  cmdOK.Top = cmdCancel.Top
  cmdOK.Left = cmdCancel.Left - (cmdOK.Width + 120)

End Function

'This Function populates a node's children and checks for grandchildren.
Private Function AddChildren(Key As String)

  Dim lFolder As Variant   'The subfolder path for the FSO
  Dim FolderKey As String  'The key of the new folder
  Dim KeyDisplay As String 'The Display Name for the node
  Dim Path As String       'The Path of Node(Key)
  
  Path = NS.Item(Key).FolderPath 'Gets the path from the clsNodeCol
  
  For Each lFolder In FSO.GetFolder(Path).SubFolders 'Looks at each subfolder
    
    'KeyDisplay = "C:\TEMP" - "C:\" = "TEMP"
    KeyDisplay = Right$(lFolder, Len(lFolder) - Len(Path))
    'In some cases KeyDisplay is preceeded by "\" so we truncate it.
    If Left$(KeyDisplay, 1) = "\" Then KeyDisplay = Right$(KeyDisplay, Len(KeyDisplay) - 1)
    'This builds the Key for the new Node
    FolderKey = Key & "\" & KeyDisplay
    'In some cases a "\\" is encountered when we only want a "\", this
    'fixes it up.
    FolderKey = Replace(FolderKey, "\\", "\")
        
    'Add the node to the tree
    DS.Nodes.Add Key, tvwChild, FolderKey, KeyDisplay, "ClosedFolder"
    'When the folder is expanded, this icon will be displayed.
    DS.Nodes(FolderKey).ExpandedImage = "OpenFolder"
    'Setup the NS collection
    NS.AddNode FolderKey
    'Set the path of the node
    NS.Item(FolderKey).FolderPath = lFolder
        
    'Check for subfolders (children) of the new node
    If FSO.GetFolder(lFolder).SubFolders.Count = 0 Then
      'If we don't find any children, we set this so that we don't
      'bother looking again.
      NS.Item(FolderKey).HasBeenBrowsed = True
    Else
      'To save time and resources, we don't want to search the
      'entire computer for all of the folders. Instead, we only search
      'the folder that we are looking at. We do need to do this search
      'so that the nodes in the tree will have a plus symbol beside
      'them so that we understand that there are subfolders.
      'Since we might not care that this subfolder exists, we simply
      'add a dummy node below that current node so that the plus appears.
      'When we select the subfolder, we will remove the dummy node and
      'populate it with the proper data.
      DS.Nodes.Add FolderKey, tvwChild, "t" & FolderKey, "dummy"
    End If
    
  Next lFolder

  'We don't sort the desktop because it mixes up the desktop folders
  'and the other icons such as My Computer. We can manually sort the desktop,
  'but we might be too lazy to do that right now :)
  If Key <> "Desktop" Then DS.Nodes(Key).Sorted = True

End Function

'This Function removes all node children
Private Function RemoveChildren(Key As String)

  Dim x As Integer
    
  'If the currnet node doesn't have any children, get out
  If DS.Nodes(Key).Children = 0 Then Exit Function

  'We count down, because as we are removing children from the top, the
  'total number of .Children goes down. If you don't get this problem,
  'change this line to:
  'For x = 1 to DS.Nodes(Key).Children
  'and see what happens.
  For x = DS.Nodes(Key).Children To 1 Step -1
    DS.Nodes.Remove DS.Nodes(Key).Child.Key
  Next x

End Function

'This Function initially populates our tree
Private Function PopulateTree()

  Dim lDrive As Variant 'The drives variable for the FSO
    
  'This next section sets up the special icons in the order that
  'we want to see them.
  
  'Add the Desktop icon
  DS.Nodes.Add , , "Desktop", "Desktop", "Desktop"
  'Add the NS for the desktop
  NS.AddNode "Desktop"
  'This gets the regisrty key for the current user's settings to display
  'the proper personalized data.
  NS.Item("Desktop").FolderPath = GetStringKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
  'Indicate that we do not want to repopulate the desktop when it expands.
  NS.Item("Desktop").HasBeenBrowsed = True
    
  'Same as above
  DS.Nodes.Add "Desktop", tvwChild, "My Documents", "My Documents", "MyDocuments"
  NS.AddNode "My Documents"
  NS.Item("My Documents").FolderPath = GetStringKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
  DS.Nodes.Add "My Documents", tvwChild, "tMy Documents", "dummy"
    
  'Same as above
  DS.Nodes.Add "Desktop", tvwChild, "My Computer", "My Computer", "MyComputer"
  NS.AddNode "My Computer"
  NS.Item("My Computer").HasBeenBrowsed = True
  NS.Item("My Computer").CanClickOK = False
  
  'I'm not quite sure how to deal with networks and since I don't have a
  'network to test this on, I will ignore it.
  'Same as above
  'DS.Nodes.Add "Desktop", tvwChild, "Network Neighbourhood", "Network Neighbourhood", "NetworkNeighbourhood"
  'NS.AddNode "Network Neighbourhood"
  'NS.Item("Network Neighbourhood").HasBeenBrowsed = True
  'NS.Item("Network Neighbourhood").CanClickOK = False
  
  'Get the data for each drive.
  For Each lDrive In FSO.Drives
    'Add the drive data to the tree
    DS.Nodes.Add "My Computer", tvwChild, lDrive & "\", lDrive, FSO.GetDrive(lDrive).DriveType + 1
    'Add the drive to NS. We need to end this in "\" because strings ending
    'in ":" sometimes create problems for keys.
    NS.AddNode lDrive & "\"
    'Set the folder for the drive
    NS.Item(lDrive & "\").FolderPath = lDrive & "\"
    'If the drive is ready (e.g. if there is a CD in CDROM) then we can
    'add more information about the drive in the tree, such as subfolders
    'and volume names.
    If FSO.Drives(lDrive).IsReady Then
      'Add a dummy node to the tree
      DS.Nodes.Add lDrive & "\", tvwChild, FSO.GetDrive(lDrive).DriveLetter, "dummy"
      'Fill in some extra data
      DS.Nodes(lDrive & "\").Text = FSO.Drives(lDrive).VolumeName & " (" & lDrive & ")"
    End If
  Next lDrive
  
  'Now we want to populate the folders of the desktop to the tree
  AddChildren "Desktop"
      
  'Expand the desktop so that we can see stuff when we start
  DS.Nodes("Desktop").Expanded = True
  
End Function


