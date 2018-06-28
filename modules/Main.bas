Attribute VB_Name = "Main"
Public Sub main()
  ' Чистим список несовпадений по строкам из результата
  ' предыдущего запуска программы (20 строк)
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  
  For Each c In activeWS.Range("B17:C107")
    c.Value = ""
  Next
  
  ' ----------------------------------------------------'
  ' Выборка исходных данных программы
  ' ----------------------------------------------------'
  
  ' Источник данных, подлежащий форматированию
  inputFilePath = activeWS.Range("C3").Value
  Dim inSrc As Workbook
  Set inSrc = Workbooks.Open(inputFilePath, True, True)
  inSheet = activeWS.Range("C4").Value
  Dim inWS As Worksheet
  Set inWS = inSrc.Worksheets(inSheet)
  
  ' Файл, в который будет осуществлена заливка из inputFile
  outputFilePath = activeWS.Range("E3").Value
  Dim outSrc As Workbook
  Set outSrc = Workbooks.Open(outputFilePath, True, True)
  outSheet = activeWS.Range("E4").Value
  Dim outWS As Worksheet
  Set outWS = outSrc.Worksheets(outSheet)
  
  ' Одинаковые по смыслу колонки из источника, заливаемого файла
  ' и кол. колонок под ними.
  inFields1 = activeWS.Range("C5").Value
  outFields1 = activeWS.Range("E5").Value
  subFields1 = activeWS.Range("C6").Value
  ' Строки, подлежащие объединению
  Dim inRows, outRows As String
  inRows = activeWS.Range("C7").Value
  outRows = activeWS.Range("E7").Value
  ' Дополнительные признаки
  inSignsStr = activeWS.Range("C8").Value
  outSignsStr = activeWS.Range("E8").Value
  ' ----------------------------------------------------'
  
  Dim inFieldMap As New Dictionary
  Dim outFieldMap As New Dictionary
  Set inFieldMap = createFieldMap((inFields1), complLevel:=(subFields1), ws:=inWS)
  Set outFieldMap = createFieldMap((outFields1), complLevel:=(subFields1), ws:=outWS)
  
  Dim unmatched As New Dictionary
  
  Dim inRowMap As New Dictionary
  Dim outRowMap As New Dictionary
  If inSignsStr <> "" And outSignsStr <> "" Then
    inSigns = Split(inSignsStr, " ")
    outSigns = Split(outSignsStr, " ")
    Set inRowMap = createRowMap((inRows), ws:=inWS, signs:=inSigns)
    Set outRowMap = createRowMap((outRows), ws:=outWS, signs:=outSigns)
    Set unmatched = mergeRowsWithSigns(inRowMap, outRowMap, inFieldMap, outFieldMap)
  Else
    Set inRowMap = createRowMap((inRows), ws:=inWS)
    Set outRowMap = createRowMap((outRows), ws:=outWS)
    Set unmatched = mergeRows(inRowMap, outRowMap, inFieldMap, outFieldMap)
  End If
  
  'i = 1
  'For Each itm In unmatched.Items
  '  r = 16 + i
  '  addr = itm.address
  '  nm = itm.name
  '  activeWS.Range("B" & r).Value = addr
  '  activeWS.Range("B" & r).Offset(0, 1).Value = nm
  '  i = i + 1
  '  Debug.Print (itm.address & " : " & itm.name)
  'Next
  
  Dim diff As New Dictionary
  Set diff = diffRows(inRowMap, outRowMap)
  Dim newRows As New Dictionary
  Dim deletedRows As New Dictionary
  Set newRows = diff.Item("new")
  Set deletedRows = diff.Item("deleted")
  
  
  
  i = 1
  activeWS.Range("B" & 17).Value = "Добавлено"
  For Each itm In newRows.Items
    r = 17 + i
    addr = itm.address
    nm = itm.name
    activeWS.Range("B" & r).Value = addr
    activeWS.Range("B" & r).Offset(0, 1).Value = nm
    i = i + 1
    'Debug.Print (itm.address & " : " & itm.name)
  Next
  
  activeWS.Range("B" & (18 + i)).Value = "Удалено"
  For Each itm In deletedRows.Items
    r = 19 + i
    addr = itm.address
    nm = itm.name
    activeWS.Range("B" & r).Value = addr
    activeWS.Range("B" & r).Offset(0, 1).Value = nm
    i = i + 1
    'Debug.Print (itm.address & " : " & itm.name)
  Next
  MsgBox ("Готово!")
End Sub
