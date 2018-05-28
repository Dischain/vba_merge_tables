Attribute VB_Name = "Main"
Public Sub main()
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  
  For Each c In activeWS.Range("B17:C37")
    c.Value = ""
  Next
  
  ' �������� ������, ���������� ��������������
  inputFilePath = activeWS.Range("C3").Value
  Dim inSrc As Workbook
  Set inSrc = Workbooks.Open(inputFilePath, True, True)
  inSheet = activeWS.Range("C4").Value
  Dim inWS As Worksheet
  Set inWS = inSrc.Worksheets(inSheet)
  
  ' ����, � ������� ����� ������������ ������� �� inputFile
  outputFilePath = activeWS.Range("E3").Value
  Dim outSrc As Workbook
  Set outSrc = Workbooks.Open(outputFilePath, True, True)
  outSheet = activeWS.Range("E4").Value
  Dim outWS As Worksheet
  Set outWS = outSrc.Worksheets(outSheet)
  
  ' ���������� �� ������ ������� �� ���������, ����������� �����
  ' � ���. ������� ��� ����.
  inFields1 = activeWS.Range("C5").Value
  outFields1 = activeWS.Range("E5").Value
  subFields1 = activeWS.Range("C6").Value
  ' ������, ���������� �����������
  Dim inRows, outRows As String
  inRows = activeWS.Range("C7").Value
  outRows = activeWS.Range("E7").Value
   
  Dim inFieldMap As New Dictionary
  Dim outFieldMap As New Dictionary
  Set inFieldMap = createFieldMap((inFields1), complLevel:=(subFields1), ws:=inWS)
  Set outFieldMap = createFieldMap((outFields1), complLevel:=(subFields1), ws:=outWS)
  
  Dim inRowMap As New Dictionary
  Dim outRowMap As New Dictionary
  Set inRowMap = createRowMap((inRows), ws:=inWS)
  Set outRowMap = createRowMap((outRows), ws:=outWS)

  'mergeSingleRow 17, 17, inFieldMap, outFieldMap
  Dim unmatched As New Dictionary
  Set unmatched = mergeRows(inRowMap, outRowMap, inFieldMap, outFieldMap)
  
  i = 1
  For Each itm In unmatched.Items
    r = 16 + i
    addr = itm.address
    nm = itm.name
    activeWS.Range("B" & r).Value = addr
    activeWS.Range("B" & r).Offset(0, 1).Value = nm
    i = i + 1
    Debug.Print (itm.address & " : " & itm.name)
  Next
  
  MsgBox ("������!")
End Sub
