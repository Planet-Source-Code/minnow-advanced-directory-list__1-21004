VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Various Properties for each node that is added to the tree

Dim m_blnHasBeenBrowsed As Boolean
Dim m_strNodeID As String
Dim m_strFolderPath As String
Dim m_blnCanClickOK As Boolean

Public Property Get CanClickOK() As Boolean

  CanClickOK = m_blnCanClickOK

End Property

Public Property Let CanClickOK(v_blnCanClickOK As Boolean)

  m_blnCanClickOK = v_blnCanClickOK

End Property

Public Property Get HasBeenBrowsed() As Boolean

  HasBeenBrowsed = m_blnHasBeenBrowsed

End Property

Public Property Let HasBeenBrowsed(v_blnHasBeenBrowsed As Boolean)

  m_blnHasBeenBrowsed = v_blnHasBeenBrowsed

End Property

Public Property Get NodeID() As String

  NodeID = m_strNodeID

End Property

Public Property Let NodeID(v_strNodeID As String)

  m_strNodeID = v_strNodeID

End Property

Public Property Get FolderPath() As String

  FolderPath = m_strFolderPath

End Property

Public Property Let FolderPath(v_strFolderPath As String)

  m_strFolderPath = v_strFolderPath

End Property

Private Sub Class_Initialize()

  HasBeenBrowsed = False
  m_blnCanClickOK = True

End Sub

