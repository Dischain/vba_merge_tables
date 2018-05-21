VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComplexField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pParent As ComplexField
Private pCell As Range
Private pNumChildren As Long
Private pChildren() As ComplexField
Private pComplexityLevel As Integer
Private pPath As String

Public Sub init(c As Range, Optional level As Integer, Optional parent As ComplexField)
  pName = c.value
  pPath = c.value
  pNumChildren = 0
  pComplexityLevel = level
  
  Set pCell = c
  Set parent = parent
  
  Me.buildSubFields level:=pComplexityLevel, initial:=Me
End Sub

Public Sub addChild(child As ComplexField)
  ReDim Preserve pChildren(pNumChildren)
  Set pChildren(pNumChildren) = child
  
  pNumChildren = pNumChildren + 1
  
  child.setParent p:=Me
End Sub

Public Function valueAt(row As Long) As Range
  Dim letter As String
  letter = Split(Cells(1, pCell.Column).address, "$")(1)
  
  Set valueAt = Range(letter & row)
End Function

Public Property Get children() As Variant
    children = pChildren
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Get address() As String
    address = pCell.address
End Property

Public Property Get parent() As ComplexField
  Set parent = pParent
End Property

Public Property Get path() As String
  path = pPath
End Property

Public Sub setParent(p As ComplexField)
  Set pParent = p
  pPath = p.path & "/" & pPath
End Sub

Public Property Get numChildren() As Long
    numChildren = pNumChildren
End Property

Public Function hasChildren() As Boolean
  hasChildren = pNumChildren <> 0
End Function

Public Sub buildSubFields(level As Integer, initial As ComplexField)
  If level > 0 Then
    initial.combineSubFields init:=initial
    If initial.numChildren <> 0 Then
      For i = 0 To UBound(initial.children)
        initial.children(i).buildSubFields (level - 1), (initial.children(i))
      Next
    End If
  End If
End Sub

Public Sub combineSubFields(init As ComplexField)
  subCols = combineSubCols(init.address)
  
  For i = 0 To UBound(subCols)
    Dim subCol As ComplexField
    Set subCol = subCols(i)
    init.addChild subCol
  Next i
End Sub

