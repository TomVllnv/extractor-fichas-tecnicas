Attribute VB_Name = "AutoajustarColumnas"
Option Explicit

' Macro OPCIONAL para Microsoft Excel de escritorio.
' El .xlsx generado no puede incluir ni ejecutar VBA automáticamente.
' Si se desea autoajuste exacto de Excel, importar el módulo y ejecutarlo.
Public Sub AutoajustarColumnas()
    Dim hoja As Worksheet
    For Each hoja In ThisWorkbook.Worksheets
        hoja.Activate
        Cells.EntireColumn.AutoFit
    Next hoja
End Sub
