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
Private pFields() As ComplexField
Private pNumFields As Long

Public Sub init(cell As Range, fieldsRange As Range, fieldsComplexityLevel As String)
  pName = cell.value
  pNumFields = 0
  
  Set pCell = cell
  
  Call addFieldsAsRange(fieldsRange, fieldsComplexityLevel)
End Sub

Public Sub addFieldsAsRange(fieldsRange As Range, fieldsComplexityLevel As String)
  For Each cell In fieldsRange
    If cell.value <> "" Then
      Dim cf As ComplexField
      Set cf = ComplexFieldFactory.CreateComplexField(Range(cell.address), (fieldsComplexityLevel))
            
      ReDim Preserve pFields(pNumFields)
      Set pFields(pNumFields) = cf
      pNumFields = pNumFields + 1
    End If
  Next
End Sub

Public Sub addFields(fields() As ComplexField)
  For Each field In fields
    pNumFields = pNumFields + 1
    ReDim Preserve pFields(pNumFields)
    Set pFields(pNumFields) = field
  Next
End Sub

Public Function getVal(fieldName As String) As String
  For i = 1 To pNumFields
    If pFields(i).name = fieldName Then
      getVal = pFields(i).valueAt(pCell.row)
      Exit For
    Else
      getVal = ""
    End If
  Next
End Function

Public Function getValByPath(p As String) As String
  For i = 0 To pNumFields - 1
    result = recursiveSearch(p, pFields(i), pCell.row)
    If result <> "" Then
      getValByPath = result
      Exit For
    Else
      getValByPath = ""
    End If
  Next
End Function

Public Function recursiveSearch(p As String, f As ComplexField, row As Long) As String
  Dim found As String
  found = ""
  rs p:=p, f:=f, row:=row, found:=found
  recursiveSearch = found
End Function

Private Function rs(p As String, f As ComplexField, row As Long, found As String) As String
  If f.path = p Then
    found = f.valueAt(row)
  Else
    If f.hasChildren() Then
      For i = 0 To f.numChildren - 1
        Dim c As ComplexField
        Set c = f.children(i)
        rs = rs(p:=p, f:=c, row:=row, found:=found)
      Next i
    Else
      rs = ""
      Exit Function
    End If
  End If
End Function
