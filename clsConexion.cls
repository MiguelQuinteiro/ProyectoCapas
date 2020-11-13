VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Hacer Referencia a
' Microsoft DAO 3.6 Object Library
' Microsoft ActiveX Data Objects 6.0 Library
Option Explicit

Public cnn As ADODB.Connection

' Conectar a la Base de Datos
Public Function Conectar(ByVal pBaseDato As String) As Boolean
  On Error GoTo Error
  Set cnn = New ADODB.Connection

  cnn.Open "Provider=SQLOLEDB.1;Password=" & "Password1" & _
           ";Persist Security Info=False;User ID=" & "sa" & _
           ";Initial Catalog=BaseImac;Data Source=" & "192.168.1.103" & _
           ";Trusted_Connection=yes"

  ''  ' Conectar a la base de datos
  ''  cnn.Open _
   ''      "Provider=sqloledb;" & _
   ''      "Data Source=LAPTOPMIGUEL\SQLEXPRESS;" & _
   ''      "Initial Catalog=NumerosPrimos;" & _
   ''      "Trusted_Connection=yes;"

  '''  ' Conectar a la base de datos
  '''  cnn.Open _
   '''      "Provider=sqloledb;" & _
   '''      "Data Source=LAPTOPMIGUEL\SQLEXPRESS;" & _
   '''      "Initial Catalog=" & pBaseDato & ";" & _
   '''      "Trusted_Connection=yes;"

  Conectar = True
  Exit Function
Error:
  MsgBox Err
  Conectar = False
End Function

' Conectar a la Base de Datos
Public Function Desconectar() As Boolean
  On Error GoTo Error
  If cnn.State = 1 Then
    cnn.Close
    Desconectar = True
  Else
    Desconectar = False
  End If
  Exit Function
Error:
  MsgBox Err
  Desconectar = False
End Function





