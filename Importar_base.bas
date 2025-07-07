Attribute VB_Name = "Importar_base"

Sub Importar_base()
'
Dim Hoja


   'Sacamos los filtros
     If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

     If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData

  'Borrar celdas de la columna B vacías
    Sheets("Export").Select
    Range("A2:AA15000").EntireRow.Delete
    On Error Resume Next
    Cells.EntireColumn.Hidden = False
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    

    Sheets("Macro").Select
    Hoja = ActiveSheet.Name
    If Hoja = "Export" Then
       MsgBox _
       ("Posicionate en otra hoja para poder ejecutar proceso"), vbExclamation _
       , "Pass ON TI"
       End
    End If
    Sheets("Export").Select
       ActiveSheet.Next.Select
       
    For Each hojas In ActiveWorkbook.Worksheets
    Hoja = ActiveSheet.Name
    'mostramos todas las columnas
     
     Cells.EntireColumn.Hidden = False
     
    'Sacamos los filtros
     If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

     If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData

    
    Range("A2").Select
    If ActiveCell.Value <> "" Then
       Cells.EntireColumn.Hidden = False
       Range("A2:AA2").Select
       Range(Selection, Selection.End(xlDown)).Select
       Selection.Copy
       
       'ocultar columnas
        Columns("I:J").EntireColumn.Hidden = True
        Columns("L:L").EntireColumn.Hidden = True
        Columns("M:P").EntireColumn.Hidden = True
        Columns("W:Z").EntireColumn.Hidden = True
       
       Sheets("Export").Select
    
    Range("A2").Select
      Do While ActiveCell.Value <> ""
           ActiveCell.Offset(1, 0).Select
      Loop
       Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
       Range("A2").Select
    
    End If
    
        Sheets(Hoja).Select
           ActiveSheet.Next.Select
        On Error GoTo Fin
        Next hojas
Fin:
        Sheets("Export").Select
           
          Do While ActiveCell.Value <> ""
               ActiveCell.Offset(1, 0).Select
          Loop
          fila = ActiveCell.Row - 2
          
          
          
          MsgBox ("Termininó el proceso, se consolidaron " & fila & " registros"), vbExclamation, "PASS ON TI"
          
    Range("a1").EntireColumn.Delete
    
    'Mostramos columnas de hojas específicas
    
     If Worksheets("Instalación de Derivativas").Activate Then
    Columns("X:Z").EntireColumn.Hidden = False
    End If
    
    If Worksheets("Hernia Laminectomia Fijacion").Activate Then
    Columns("M:O").EntireColumn.Hidden = False
    End If
    If Worksheets("Cesárea cs salpingoligadu").Activate Then
    Columns("P").EntireColumn.Hidden = False
    Columns("W").EntireColumn.Hidden = False


    End If
    
    Worksheets("Export").Activate
    
    
    End Sub
    


