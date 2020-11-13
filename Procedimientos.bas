Attribute VB_Name = "Procedimientos"
' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Public Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
  If Not o_Excel Is Nothing Then Set o_Excel = Nothing
  If Not o_Libro Is Nothing Then Set o_Libro = Nothing
  If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub
