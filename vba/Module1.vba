Sub ExportGrid()
    On Error GoTo ErrHandler
    Dim gridExporter As Object
    Set gridExporter = CreateObject("RaghavStaadExporter.Grid")
    gridExporter.ExportGridFromsheet4
    Exit Sub
ErrHandler:
    MsgBox "Error while exporting grid: " & Err.Description, vbExclamation, "DXF Export"
End Sub

