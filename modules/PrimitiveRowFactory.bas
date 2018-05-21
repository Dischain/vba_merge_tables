Attribute VB_Name = "PrimitiveRowFactory"
Public Function CreatePrimitiveRow(cl As Range, fr As Range, fcl As Integer) As PrimitiveRow
  Dim pr_obj As PrimitiveRow
  Set pr_obj = New PrimitiveRow
  
  pr_obj.init cell:=cl, fieldsRange:=fr, fieldsComplexityLevel:=(fcl)
  
  Set CreatePrimitiveRow = pr_obj
End Function
