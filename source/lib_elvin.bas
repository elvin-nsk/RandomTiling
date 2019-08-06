Attribute VB_Name = "lib_elvin"
'=======================================================================================
' Модуль:            lib_elvin
' Версия:            19.08.05
' Автор:             elvin-nsk (me@elvin.nsk.ru)
' Использован код:   dizzy (из макроса CtC)
'                    и др.
' Назначение:        библиотека функций для макросов от elvin-nsk
' Использование:
' Зависимости:       самодостаточный
'=======================================================================================

Option Explicit


'---------------------------------------------------------------------------------------
' Функции          : boostStart, boostFinish
' Версия           : 19.08.05
' Авторы           : dizzy, elvin-nsk
' Назначение       : доработанные оптимизаторы от CtC
' Зависимости      : самодостаточные
'
' Параметры:
' ~~~~~~~~~~
'
'
' Использование:
' ~~~~~~~~~~~~~~
'
'---------------------------------------------------------------------------------------
Public Sub boostStart(Optional ByVal unDo$ = "", Optional ByVal optimize = True)
  If unDo <> "" And Not (ActiveDocument Is Nothing) Then ActiveDocument.BeginCommandGroup unDo
  If optimize Then Optimization = True
  EventsEnabled = False
  If Not ActiveDocument Is Nothing Then
    With ActiveDocument
      .SaveSettings
      '.PreserveSelection = False отключено, вызывает глюки с intersect, на производительность при включенной оптимизации почти не влияет
      .Unit = cdrMillimeter
      .ReferencePoint = cdrCenter
    End With
  End If
End Sub
Public Sub boostFinish(Optional ByVal endUndoGroup = False)
  EventsEnabled = True
  Optimization = False
  If Not (ActiveDocument Is Nothing) Then
    With ActiveDocument
      .RestoreSettings
      If endUndoGroup Then .EndCommandGroup
    End With
    ActiveWindow.Refresh
  End If
  Application.Refresh
End Sub


Function FindShapesByName(SourceRange As ShapeRange, ByVal Name$) As ShapeRange
  Set FindShapesByName = SourceRange.Shapes.findshapes(Name)
End Function

Function FindShapesByNamePart(SourceRange As ShapeRange, ByVal NamePart$) As ShapeRange
  Set FindShapesByNamePart = SourceRange.Shapes.findshapes(Query:="@name.contains('" & NamePart & "')")
End Function


'функция отсюда: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
Public Function isStrInArr(ByVal stringToBeFound$, arr As Variant) As Boolean
    Dim i&
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            isStrInArr = True
            Exit Function
        End If
    Next i
    isStrInArr = False
End Function

'Returns the rightmost characters of a string upto but not including the rightmost '\'
'e.g. 'c:\winnt\win.ini' returns 'win.ini'
'не моё :)
Public Function getFilenameFromPath(ByVal strPath$) As String
  If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
    getFilenameFromPath = getFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
  End If
End Function

Public Function getFolderFromPath(ByVal strPath$) As String
  If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
    getFolderFromPath = Left(strPath, InStrRev(strPath, "\"))
  End If
End Function

'заменяет расширение файлу на заданное
Function SetFileExt(ByVal SourceFile, ByVal NewExt) As String

  Dim pos&, ret$
  
  pos = InStr(SourceFile, ".")
  If pos > 0 Then
    ret = Left(SourceFile, pos) & NewExt
  Else
    ret = SourceFile
  End If
  
  SetFileExt = ret

End Function

'является ли число чётным :) Что такое Even и Odd запоминать лень...
Public Function isChet(ByVal X) As Boolean
  If X Mod 2 = 0 Then isChet = True Else isChet = False
End Function

