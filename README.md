# glowing-fishstick
TALLER DE PRODUCTIVIDAD BASADA EN HERRAMIENTAS TECNOLÓGICAS
Código para Automatizar el Reporte de Artículos Comprometidos  
Sub Macro3()
'
' Macro3 Macro
'
'LIMPIAR DATOS DE ESTA TABLA'

    Application.Goto Reference:="R9C1"
    Rows("9:9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Application.Goto Reference:="R9C1"
    
'FIN DE LIMPIAR DATOS DE ESTA TABLA'

    ChDir "C:\Francisco\REPORTES\suministro"
          Workbooks.OpenText Filename:="C:\Francisco\REPORTES\suministro\suministro.txt" _
        , Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier _
        :=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:= _
        False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1) _
        , TrailingMinusNumbers:=True
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "aaa"
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$50000").AutoFilter Field:=1, Criteria1:=Array( _
        "---------------------------------------------------------------------------------------------------------------------------------------" _
        , "
", _
        "_______________________________________________________________________________________________________________________________________" _
        , _
        "GPO GEN ESP  DI VA U.M CANT. PRES T.P DESCRIPCION                                                   No. ORDEN        CANTIDAD EDO" _
        , "Reporte general de art¡culos comprometidos o en embarque ordenado por unidad" _
        ), Operator:=xlFilterValues
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "aaa"
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$50000").AutoFilter Field:=1, Criteria1:= _
        "=Clas.*", Operator:=xlOr, Criteria2:="=Unidad:*"
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "aaa"
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$50000").AutoFilter Field:=1, Criteria1:= _
        "=CONTROL*", Operator:=xlAnd
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$50000").AutoFilter Field:=1, Criteria1:= _
        "=CLAS_PTAL*", Operator:=xlAnd
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Range("C1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Range("E1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$50000").AutoFilter Field:=1, Criteria1:= _
        "=??? ??? ???? ?? ??*", Operator:=xlAnd
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.ClearContents
    Range("C1").Select
    Selection.CurrentRegion.Select
    Selection.AutoFilter
    ActiveSheet.Range("$C$1:$C$50000").AutoFilter Field:=1, Criteria1:="="
    Range("C1").Select
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Range("A50010").Select
    ActiveCell.FormulaR1C1 = "fin"
    Range("A1").Select
    
    Do
    Selection.End(xlDown).Select
    If ActiveCell = "fin" Then Exit Do
    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.End(xlDown).Select
    Loop
   
    Range("C50010").Select
    ActiveCell.FormulaR1C1 = "fin"
    Range("C50013").Select
    ActiveCell.FormulaR1C1 = "fin"
    Range("C1").Select
    Selection.Copy
    Selection.CurrentRegion.Select
    ActiveSheet.Paste
    
    Do
    Selection.End(xlDown).Select
    If ActiveCell = "fin" Then Exit Do
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.CurrentRegion.Select
    ActiveSheet.Paste
    Loop

    Range("A1").Select
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "a"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "a"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "a"
    Range("A1:J50000").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$50000").AutoFilter Field:=1, Criteria1:="<>"
    Range("A2:A50000").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$E$50000").AutoFilter Field:=1
    ActiveSheet.Range("$A$1:$E$50000").AutoFilter Field:=3, Criteria1:="="
    Range("C1:C50000").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Columns("C:C").Select
    Selection.Cut
    Columns("A:A").Select
    ActiveSheet.Paste
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        OtherChar:=":", FieldInfo:=Array(Array(0, 2), Array(12, 1)), _
        TrailingMinusNumbers:=True
    Range("E1").Select
    Selection.CurrentRegion.Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlFixedWidth, _
        OtherChar:=":", FieldInfo:=Array(Array(0, 2), Array(3, 2), Array(18, 1), Array(22 _
        , 1), Array(33, 1), Array(37, 1), Array(100, 1), Array(109, 1), Array(125, 1)), _
        TrailingMinusNumbers:=True
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Cut
    Windows("Macro Remisiones.xlsm").Activate
    Range("A9").Select
    ActiveSheet.Paste
    Windows("suministro.txt").Activate
    Range("F1").Select
    Selection.CurrentRegion.Select
    Selection.Cut
    Windows("Macro Remisiones.xlsm").Activate
    Range("E9").Select
    ActiveSheet.Paste
    
    Range("A1").Select
    Windows("suministro.txt").Activate
    ActiveWindow.Close
    
'CREAR ARCHIVO PENDIENTE

    ChDir "C:\Francisco\REPORTES\suministro"
    Workbooks.Open Filename:="C:\Francisco\REPORTES\suministro\PENDIENTE.xlsx"
    Application.Goto Reference:="R2C1"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Windows("Macro Remisiones.xlsm").Activate
    Application.Goto Reference:="R9C1"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("PENDIENTE.xlsx").Activate
    Application.Goto Reference:="R2C1"
    ActiveSheet.Paste
    ActiveWindow.SmallScroll ToRight:=5
    Range("K2").Select
    Application.CutCopyMode = False
    
'    ActiveWorkbook.Worksheets("suministro").AutoFilter.Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("suministro").AutoFilter.Sort.SortFields.Add Key:= _
        Range ("K1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
'   With ActiveWorkbook.Worksheets("suministro").AutoFilter.Sort
'       .Header = xlYes
'       .MatchCase = False
'       .Orientation = xlTopToBottom
'       .SortMethod = xlPinYin
'       .Apply
'  End With
    
    ActiveWorkbook.Save
    
    'FIN DE CREAR ARCHIVO PENDIENTE
    Range("K2").Select

End Sub
