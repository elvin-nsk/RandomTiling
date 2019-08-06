Attribute VB_Name = "RandomTiling"
'=======================================================================================
' Макрос:            Случайное замощение (elvin_RandomTiling)
' Версия:            19.08.06
' Автор:             elvin-nsk (me@elvin.nsk.ru)
' Назначение:        Случайное замощение выделенными объектами заданной площади.
'                    Все размеры объектов приводятся к заданному.
'=======================================================================================

Option Explicit

Sub start()
  
  Dim e As Shape, elements As ShapeRange
  Dim ew#, eh#, en&, StartX#, StartY#
  Dim rows&, cols&, row&, col&, rot As Boolean, del As Boolean
  
  If ActiveSelection.Shapes.Count = 0 Then Exit Sub
  
  Set elements = ActiveSelection.Shapes.All
  ActiveDocument.Unit = cdrMillimeter
  
  With frm_Dialog
    .SelNum = "Выделено " & elements.Count & " элементов"
    .ElementW = CStr(elements.FirstShape.SizeWidth)
    .ElementH = CStr(elements.FirstShape.SizeHeight)
    .RowsNum = CStr(10)
    .ColsNum = CStr(10)
    .Show
    ew = CDbl(.ElementW)
    eh = CDbl(.ElementH)
    rows = CLng(.RowsNum)
    cols = CLng(.ColsNum)
    del = CBool(.cbDelete)
    rot = CBool(.cbRotate)
    If .isOk = False Then Exit Sub
  End With
  
  boostStart "Случайное замощение", True
  
  If rot Then eh = ew 'если вращаем, то стороны должны быть одинаковые
  StartX = 0
  StartY = 0
  Randomize
  For row = 1 To rows
    For col = 1 To cols
      en = Int((elements.Count * Rnd) + 1)
      Set e = elements(en).Duplicate
      e.SizeWidth = ew
      e.SizeHeight = eh
      e.LeftX = StartX + (col - 1) * ew
      e.BottomY = StartY + (row - 1) * eh
      If rot Then e.Rotate ((Int((3 * Rnd) + 1)) * 90)
    Next
  Next
  
  If del Then elements.Delete

  boostFinish True
  
End Sub
