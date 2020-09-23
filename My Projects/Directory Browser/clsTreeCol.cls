VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNodeCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colNode As Collection

Public Property Get Count() As Long

  'Returns the total number of nodes in the collection
  Count = m_colNode.Count

End Property

Public Sub AddNode(ByVal v_strNodeID As String)
  
  'Setup a temp clsNode
  Dim objNode As New clsNode
  
  'If no Key was passed then exit the sub
  If v_strNodeID = "" Then
    Exit Sub
  End If
    
  objNode.NodeID = v_strNodeID
  'Add the node to the collection
  m_colNode.Add objNode, objNode.NodeID
  
  'Clear the temp node
  Set objNode = Nothing
 
End Sub

Public Sub Clear()

  Dim x As Integer
  
  'If the collection is empty, get out
  If m_colNode.Count = 0 Then Exit Sub
  
  'Clear out the collection, one at a time
  For x = m_colNode.Count To 1 Step -1
    m_colNode.Remove x
  Next x
  
End Sub

Public Function Item(ByVal v_vntIndex As Variant) As clsNode

  On Error Resume Next
  'Attempt to set the item to the collection item
  Set Item = m_colNode.Item(v_vntIndex)
  
  'If there is an error, return nothing
  If Err.Number <> 0 Then
    Set Item = Nothing
  End If
  
End Function

Private Sub Class_Initialize()

  Set m_colNode = New Collection
  
End Sub



