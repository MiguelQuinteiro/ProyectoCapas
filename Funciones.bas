Attribute VB_Name = "Funciones"
Option Explicit

' -------------------------------------------------------------------------------------------
' \\ -- Función para crear un nuevo libro con el contenido del Grid
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean

  On Error GoTo Error_Handler

  Dim o_Excel As Object
  Dim o_Libro As Object
  Dim o_Hoja As Object
  Dim Fila As Long
  Dim Columna As Long

  ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
  Set o_Excel = CreateObject("Excel.Application")
  Set o_Libro = o_Excel.Workbooks.Add
  Set o_Hoja = o_Libro.Worksheets.Add

  ' -- Bucle para Exportar los datos
  With FlexGrid
    For Fila = 0 To .Rows - 1
      For Columna = 0 To .Cols - 1
        If Fila = 0 Then
          o_Hoja.Cells(Fila + 1, Columna + 1).Value = .TextMatrix(Fila, Columna)
        Else
          o_Hoja.Cells(Fila + 1, Columna + 1).Value = .TextMatrix(Fila, Columna)
        End If
      Next
    Next
  End With

  o_Libro.Close True, sOutputPath
  ' -- Cerrar Excel
  o_Excel.Quit
  ' -- Terminar instancias
  Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
  Exportar_Excel = True
  Exit Function

  ' -- Controlador de Errores
Error_Handler:
  ' -- Cierra la hoja y el la aplicación Excel
  If Not o_Libro Is Nothing Then: o_Libro.Close False
  If Not o_Excel Is Nothing Then: o_Excel.Quit
  Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
  If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function


