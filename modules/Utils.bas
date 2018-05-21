Attribute VB_Name = "Utils"
Public Function combineSubCols(addr As String) As Variant
  Dim parentColl, currentSub, parentSibling As Range
  Dim length As Integer
  Dim combined() As ComplexField
  
  Set parentColl = Range(addr)
  Set parentSibling = parentColl.Offset(0, 1)
  Set currentSub = parentColl.Offset(1, 0)
  
  length = 0
  
  Do While currentSub.Column <> parentSibling.Column
    Dim c As ComplexField
    Set c = ComplexFieldFactory.CreateComplexField(Range(currentSub.address))
    
    ReDim Preserve combined(length)
    Set combined(length) = c
    length = length + 1
    Set currentSub = currentSub.Offset(0, 1)
  Loop
  
  combineSubCols = combined
End Function
