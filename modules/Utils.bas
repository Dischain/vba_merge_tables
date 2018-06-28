Attribute VB_Name = "Utils"
Public Function combineSubCols(addr As String, ws As Worksheet) As Variant
  Dim parentColl, currentSub, parentSibling As Range
  Dim length As Integer
  Dim combined() As ComplexField
  
  Set parentColl = ws.Range(addr)
  Set parentSibling = parentColl.Offset(0, 1)
  Set currentSub = parentColl.Offset(1, 0)
  
  length = 0
  
  Do While currentSub.Column <> parentSibling.Column
    Dim c As ComplexField
    Set c = ComplexFieldFactory.CreateComplexField(ws, currentSub.address)
    
    ReDim Preserve combined(length)
    Set combined(length) = c
    length = length + 1
    Set currentSub = currentSub.Offset(0, 1)
  Loop
  
  combineSubCols = combined
End Function

Public Function createFieldMap(addr As String, complLevel As Integer, ws As Worksheet) As Dictionary
  Dim fieldMap As Dictionary
  Set fieldMap = New Dictionary
  
  For Each f In ws.Range(addr)
    If f.Value <> "" Then
      Dim field As ComplexField
      Set field = ComplexFieldFactory.CreateComplexField(ws, (f.address), l:=complLevel)
      
      fieldMap.Add Key:=field.name, Item:=field
    End If
  Next
  
  Set createFieldMap = fieldMap
End Function

Public Function createRowMap(addr As String, ws As Worksheet, Optional signs As Variant) As Dictionary
  Dim rowMap As Dictionary
  Set rowMap = New Dictionary
  
  For Each r In ws.Range(addr)
    If r.Value <> "" And Not containsEscapeWords(r.Value) Then
      Dim row As PrimitiveRow
      Set row = PrimitiveRowFactory.CreatePrimitiveRow(ws, (r.address), signs)
            
      rowMap.Add Key:=row.name, Item:=row
    End If
  Next
  
  Set createRowMap = rowMap
End Function

' Выполняет слияние множества строк с предварительной проверкой совпадения по именам строк, без учета доп. признаков
Public Function mergeRows(inRowMap As Dictionary, outRowMap As Dictionary, inFieldMap As Dictionary, outFieldMap As Dictionary) As Dictionary
  Dim notMatched As Dictionary
  Set notMatched = New Dictionary

  For Each inRow In inRowMap.Items
    If outRowMap.Exists(inRow.name) Then
      Dim outRow As PrimitiveRow
      Set outRow = outRowMap.Item(inRow.name)
      mergeSingleRow inRow:=inRow.row, outRow:=outRow.row, inFieldMap:=inFieldMap, outFieldMap:=outFieldMap
    Else
      notMatched.Add Key:=inRow.name, Item:=inRow
    End If
  Next

  Set mergeRows = notMatched
End Function

' Выполняет слияние множества строк с предварительной проверкой совпадения по именам строк, с учетом доп. признаков
Public Function mergeRowsWithSigns(inRowMap As Dictionary, outRowMap As Dictionary, inFieldMap As Dictionary, outFieldMap As Dictionary) As Dictionary
  Dim notMatched As Dictionary
  Set notMatched = New Dictionary

  For Each inRow In inRowMap.Items
    If outRowMap.Exists(inRow.name) Then
      Dim outRow As PrimitiveRow
      Set outRow = outRowMap.Item(inRow.name)
      
      Dim inRowSigns As New Dictionary
      Dim outRowSigns As New Dictionary
      Set inRowSigns = inRow.signs
      Set outRowSigns = outRow.signs
      
      If dictEquals(inRowSigns, outRowSigns) Then
        mergeSingleRow inRow:=inRow.row, outRow:=outRow.row, inFieldMap:=inFieldMap, outFieldMap:=outFieldMap
      Else
        notMatched.Add Key:=inRow.name, Item:=inRow
      End If
    Else
      notMatched.Add Key:=inRow.name, Item:=inRow
    End If
  Next

  Set mergeRowsWithSigns = notMatched
End Function

' Выполняет слияние двух строк путем проверки всех полей
Public Sub mergeSingleRow(inRow As Long, outRow As Long, inFieldMap As Dictionary, outFieldMap As Dictionary)
  For Each inField In inFieldMap.Items
    If outFieldMap.Exists(Key:=inField.name) Then
      Dim outField As ComplexField
      Set outField = outFieldMap.Item(inField.name)
      
      Dim inFieldLowestFields As New Dictionary
      Set inFieldLowestFields = inField.lowestFields
      Dim outFieldLowestFields As New Dictionary
      Set outFieldLowestFields = outField.lowestFields
      
      For Each inLF In inFieldLowestFields.Items
        
        If outFieldLowestFields.Exists(Key:=inLF.path) Then
          Dim outLF As PrimitiveField
          Set outLF = outFieldLowestFields.Item(Key:=inLF.path)
                    
          inVal = inLF.getValueAt(inRow)
          outLF.setValueAt (inVal), (outRow)
        End If
      Next
    End If
  Next
End Sub

' Выполняет операцию равенства по ключам для двух хэш-таблиц
Public Function dictEquals(dict1 As Dictionary, dict2 As Dictionary) As Boolean
  For Each k In dict1.Keys
    If dict1.Item(k) <> dict2.Item(k) Then
      dictEquals = False
      Exit Function
    End If
  Next
  dictEquals = True
End Function

Public Function containsEscapeWords(name As String) As Boolean
  For Each w In escapeWords
    word = LCase(w)
    Dim str As String
    str = LCase(name)
    If startsWith((word), str) Then
      containsEscapeWords = True
      Exit Function
    End If
  Next
  containsEscapeWords = False
End Function

Public Function startsWith(s As String, seed As String) As Boolean
  If InStr(1, seed, s) = 1 Then
    startsWith = True
  Else
    startsWith = False
  End If
End Function

Public Function escapeWords() As Variant
  escapeWords = Array("Министерство", "Дирекция", "Объекты", "Модернизация", "Служба", "Государственный комитет", "Управление")
End Function

Public Function arrayToString(arr As Variant) As String
  res = ""
  
  For Each itm In arr
    res = res & itm & Chr(13)
  Next
  arrayToString = res
End Function

Public Function stringToArray(str As String) As Variant
  arr = Split(str, ",")
  Dim res() As String
  For i = 0 To UBound(arr)
    ReDim Preserve res(i + 1)
    s = Trim(arr(i))
    res(i) = s
  Next
  stringToArray = res
End Function

Public Function eraseEOLs(s As String) As String
  tempStr = ""
  For c = 1 To Len(s)
    If Mid(s, c, 1) = vbCr Or Mid(s, c, 1) = vbLf Then
      tempStr = tempStr + ""
    Else
      tempStr = tempStr + Mid(s, c, 1)
    End If
  Next
  eraseEOLs = tempStr
End Function

Public Function eraseSPs(s As String) As String
  tempStr = ""
  For c = 2 To Len(s)
    Dim prev As Integer
    prev = c - 1
    
    If c > 2 And Mid(s, c, 1) = " " And Mid(s, prev, 1) = " " Then
      tempStr = tempStr + ""
    Else
      tempStr = tempStr + Mid(s, c, 1)
    End If
  Next
  eraseSPs = Mid(s, 1, 1) & tempStr
End Function

Public Function eraseTrailingPeriod(s As String) As String
  tempStr = ""
  l = Len(s)
  For c = 1 To l
    If c = l And Mid(s, c, 1) = "." Then
      tempStr = tempStr + ""
    Else
      tempStr = tempStr + Mid(s, c, 1)
    End If
  Next
  eraseTrailingPeriod = tempStr
End Function

Public Function concat(arr1 As Variant, arr2 As Variant) As Variant
  arr1Length = UBound(arr1)
  arr2Length = UBound(arr2)
  For i = 0 To arr2Length
    ReDim Preserve arr1(arr1Length + i + 1)
    Set arr1(arr1Length + i) = arr2(i)
  Next i
  concat = arr1
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Public Function diffRows(inRowMap As Dictionary, outRowMap As Dictionary) As Dictionary
  Dim newRows As New Dictionary
  Dim deletedRows As New Dictionary
  Dim result As New Dictionary

  For Each inr In inRowMap.Items
    If Not outRowMap.Exists(inr.name) Then
      newRows.Add Key:=inr.name, Item:=inr
    End If
  Next

  For Each outr In outRowMap.Items
    If Not inRowMap.Exists(outr.name) Then
      deletedRows.Add Key:=outr.name, Item:=outr
    End If
  Next

  result.Add Key:="new", Item:=newRows
  result.Add Key:="deleted", Item:=deletedRows
  
  Set diffRows = result
End Function

