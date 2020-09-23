Attribute VB_Name = "modShowTree"
Option Explicit

Public Sub ShowTree(TreeNode As Object, FormBox As Object, ModBox As Object, _
    ClsBox As Object, ControlBox As Object, PageBox As Object, DepBox As Object, Title As String)
    Dim x As Integer
    ' Add the root of the Treeview
    addRoot Title, TreeNode, 6
    
    ' Add Section Forms
    addSec "Forms", TreeNode, 4
    ' Load form names and path into tree
    For x = 0 To (FormBox.ListCount - 1)
        addKey FormBox.List(x), TreeNode, 4
    Next x
    
    ' Add Section Modules
    addSec "Modules/BAS", TreeNode, 1
    ' Load Module names and path into tree
    For x = 0 To (ModBox.ListCount - 1)
        addKey ModBox.List(x), TreeNode, 1
    Next x
    
    ' Add Section Classes
    addSec "Classes", TreeNode, 2
    For x = 0 To (ClsBox.ListCount - 1)
        addKey ClsBox.List(x), TreeNode, 2
    Next x
    
    ' Add Section UserControls
    addSec "UserControls", TreeNode, 7
    For x = 0 To (ControlBox.ListCount - 1)
        addKey ControlBox.List(x), TreeNode, 7
    Next x
    
    ' Add Section Property Pages
    addSec "PropertyPages", TreeNode, 8
    For x = 0 To (PageBox.ListCount - 1)
        addKey PageBox.List(x), TreeNode, 8
    Next x
    
    ' Add Section Dependencies
    addSec "Dependencies", TreeNode, 9
    For x = 0 To (DepBox.ListCount - 1)
        addKey DepBox.List(x), TreeNode, 9
    Next x
    
    nodRoot.Expanded = True
End Sub
