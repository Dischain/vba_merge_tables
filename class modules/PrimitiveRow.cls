VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrimitiveRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pCell As Range
Private ws As Worksheet

Public Sub init(sheet As Worksheet, c As String)
  Set ws = sheet
  Set pCell = ws.Range(c)
  
  pName = Trim(eraseSPs(eraseEOLs(s:=pCell.Value)))
End Sub

Public Property Get row() As Long
  row = pCell.row
End Property

Public Property Get name() As String
  name = pName
End Property

Public Property Get address() As String
  address = pCell.address
End Property
