Attribute VB_Name = "Test"
Private Sub ComplexFieldTest()
  Dim field1 As ComplexField
  Set field1 = ComplexFieldFactory.CreateComplexField(Range("c6"))
  Dim field2 As ComplexField
  Set field2 = ComplexFieldFactory.CreateComplexField(Range("c8"))
  field1.addChild child:=field2
  Debug.Print (field1.valueAt(100))
  Debug.Print (field1.hasChildren)
  Debug.Print (field2.parent.name)
End Sub

Private Sub primitiveRowTest()
  Dim r1 As PrimitiveRow
  Set r1 = PrimitiveRowFactory.CreatePrimitiveRow(Range("c14"), Range("r6:aa6"), 1)
  r1.addFieldsAsRange fieldsRange:=Range("AC8:CL8"), fieldsComplexityLevel:=(3)
  
  Debug.Print ("Найдено2: " & r1.getValByPath("март/Всего/факт"))
  Debug.Print ("Найдено3: " & r1.getValByPath("февраль"))
  
  Dim testField As ComplexField
  Set testField = ComplexFieldFactory.CreateComplexField(Range("AU8"), 3)
  Debug.Print ("-------------------------")
  Debug.Print ("result: " & r1.recursiveSearch("май/Всего/план", testField, 15))
End Sub

Private Sub combineSubColsTest()
  Dim field As ComplexField
  Set field = ComplexFieldFactory.CreateComplexField(Range("AC8"), 3)

  Debug.Print ("-----")
  Debug.Print ("name: " & field.name)
  Debug.Print ("field.numChildren: " & field.numChildren)
  field.buildSubFields level:=3, initial:=field
  Dim l1 As ComplexField
  Set l1 = field.children(1)
  Dim l2 As ComplexField
  Set l2 = l1.children(0)
  Debug.Print (l2.children(0).parent.parent.name)
  Debug.Print (l2.children(0).path)
End Sub
