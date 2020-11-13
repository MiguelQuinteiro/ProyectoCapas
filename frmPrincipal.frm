VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   " .:. MANTENIMIENTO DE TABLA .:. "
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   13950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEdicion 
      Caption         =   " .:. Edición de Campos .:. "
      Height          =   3375
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxEdicionCampos 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   3
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraBaseDatos 
      Caption         =   " .:. Base de Datos .:. "
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   13695
      Begin MSFlexGridLib.MSFlexGrid flxDatos 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   7646
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraParametros 
      Caption         =   ".:. Parametros .:.  "
      Height          =   3375
      Left            =   8160
      TabIndex        =   18
      Top             =   4920
      Width           =   2775
      Begin VB.ComboBox cboBaseDato 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":0000
         Left            =   120
         List            =   "frmPrincipal.frx":000D
         TabIndex        =   3
         Text            =   "BaseImac"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cboTabla 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":0039
         Left            =   120
         List            =   "frmPrincipal.frx":0055
         TabIndex        =   4
         Text            =   "Rop_Prendas"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cboCampos 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtRegistros 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Text            =   "50"
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblCampos 
         Caption         =   "Campos"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lblTabla 
         Caption         =   "Tabla"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblBaseDatos 
         Caption         =   "Base de Datos"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblCantidad 
         Caption         =   "Cantidad de Registros"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.Frame fraBotones 
      Caption         =   " .:. Controles .:. "
      Height          =   3375
      Left            =   11040
      TabIndex        =   17
      Top             =   4920
      Width           =   2775
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Excel"
         Height          =   495
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Conectar"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   1440
         TabIndex        =   16
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdReportar 
         Caption         =   "Reportar"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   495
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdDesactivar 
         Caption         =   "Desactivar"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdListar 
         Caption         =   "Listar"
         Height          =   495
         Left            =   1440
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Label lblAviso 
      Caption         =   "ESPERE MIENTRAS SE GENERA EL ARCHIVO CON LOS DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4680
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim miCadena As String

' Boton Abrir
Private Sub cmdAbrir_Click()
  On Error GoTo Error
  ' Declaración de variables
  Dim tField As ADODB.Field
  Dim miCantidadCampos As Integer
  Dim col As Long
  Dim miFil As Long
  Dim miAnchoMinimo As Integer
  Dim miAnchoMaximo As Integer
  Dim miAnchoDato As Integer
  Dim miAnchoColumna() As Integer
  Dim i As Integer
  Dim miIncremento As Integer
  Dim Row As Long

  ' Declaración de variables y conexión
  Dim rst As ADODB.Recordset
  Set rst = New ADODB.Recordset
  Dim oConexion As New clsConexion
  Set oConexion = New clsConexion

  ' Muestra el Aviso
  flxDatos.Visible = False
  fraBaseDatos.Visible = False
  lblAviso.Visible = True
  DoEvents

  ' Conecta a la Base de Datos
  oConexion.Conectar (cboBaseDato.Text)

  ' Abrir el recordset indicando la tabla a la que queremos acceder
  rst.Open "SELECT * FROM " & cboTabla.Text, oConexion.cnn, adOpenDynamic, adLockOptimistic

  ' Inicializa
  miIncremento = 140
  miAnchoMinimo = 5

  ' Asignar los nombres de los campos al combo
  With cboCampos
    .Clear
    miAnchoMaximo = miAnchoMinimo
    For Each tField In rst.Fields
      .AddItem tField.Name
      If miAnchoMaximo < Len(tField.Name) Then
        miAnchoMaximo = Len(tField.Name)
      End If
    Next
    miAnchoMaximo = miAnchoMaximo * miIncremento
    .ListIndex = 0
  End With
  miCantidadCampos = cboCampos.ListCount

  ' Redimensiona e inicializa a cero control de ancho de columnas
  ReDim miAnchoColumna(miCantidadCampos)
  i = 0
  For Each tField In rst.Fields
    miAnchoColumna(i) = miAnchoMinimo
    i = i + 1
  Next

  ' Setear las grillas de Datos y Edición de Campos
  flxDatos.Clear
  flxDatos.FixedCols = 0
  flxDatos.Cols = miCantidadCampos
  For i = 1 To miCantidadCampos - 1
    flxDatos.ColAlignment(i) = flexAlignLeftCenter
  Next

  flxEdicionCampos.ColAlignment(1) = flexAlignLeftCenter
  flxEdicionCampos.Clear
  flxEdicionCampos.FixedRows = 1
  flxEdicionCampos.Rows = miCantidadCampos + 1
  flxEdicionCampos.ColWidth(2) = 1

  ' Determina el ancho de los títulos de las columnas
  flxDatos.Row = 0
  col = 0
  For Each tField In rst.Fields
    flxDatos.col = col
    flxDatos.Text = tField.Name
    If miAnchoColumna(col) < Len(tField.Name) Then
      miAnchoColumna(col) = Len(tField.Name)
    End If
    col = col + 1
  Next

  ' Determina el ancho de los títulos en las filas
  flxEdicionCampos.ColWidth(0) = miAnchoMaximo
  flxEdicionCampos.col = 0

  ' Coloca los nombres de los campos en las filas
  flxEdicionCampos.Row = 0
  flxEdicionCampos.col = 0
  flxEdicionCampos.Text = "Campos"
  flxEdicionCampos.col = 1
  flxEdicionCampos.Text = "Datos"
  Row = 1
  flxEdicionCampos.col = 0
  For Each tField In rst.Fields
    flxEdicionCampos.Row = Row
    flxEdicionCampos.Text = tField.Name
    Row = Row + 1
  Next
  rst.MoveFirst
  miFil = 2

  ' Coloca los datos dentro de la grilla de datos
  miAnchoDato = miAnchoMinimo
  If Not (rst.BOF And rst.EOF) Then
    While Not (rst.EOF) And miFil < Val(txtRegistros) + 2
      flxDatos.Rows = miFil + 1
      'flxDatos.Row = 1
      col = 0
      For Each tField In rst.Fields
        flxDatos.col = col
        If IsNull(rst(col)) Then
          flxDatos.Text = ""
        Else
          If IsNumeric(rst(col)) Then
            flxDatos.Text = Trim(CStr(rst(col)))
          Else
            flxDatos.Text = Trim(rst(col))
          End If
        End If

        If miAnchoDato < Len(flxDatos.Text) Then
          miAnchoDato = Len(flxDatos.Text)
        End If

        If miAnchoColumna(col) < Len(flxDatos.Text) Then
          miAnchoColumna(col) = Len(flxDatos.Text)
        End If

        col = col + 1
      Next

      rst.MoveNext
      miFil = miFil + 1
      flxDatos.Row = flxDatos.Row + 1
    Wend

    ' Quita la ultima fila
    flxDatos.RemoveItem (flxDatos.Row)
  End If

  flxDatos.Refresh
  fraBaseDatos.Visible = True
  flxDatos.Visible = True
  lblAviso.Visible = False

  col = 0
  For Each tField In rst.Fields
    flxDatos.ColWidth(col) = miAnchoColumna(col) * miIncremento
    col = col + 1
  Next

  flxEdicionCampos.ColWidth(1) = miAnchoDato * miIncremento
  DoEvents

  ' Cerrar el recordset
  rst.Close
  ' Desconecta a la Base de Datos
  oConexion.Desconectar

Error:


End Sub

' Botón de Exportar
Private Sub cmdExportar_Click()
  If Exportar_Excel(App.Path & "\ConsultaResumida.xls", flxDatos) Then
    MsgBox " Datos exportados en " & App.Path, vbInformation
  End If
End Sub

' Botón de Agregar
Private Sub cmdAgregar_Click()

End Sub

' Botón de Consultar
Private Sub cmdConsultar_Click()

End Sub

' Botón de Modificar
Private Sub cmdModificar_Click()

End Sub

' Botón de Eliminar
Private Sub cmdEliminar_Click()

End Sub

' Botón de Desactivar
Private Sub cmdDesactivar_Click()

End Sub

' Botón de Reportar
Private Sub cmdReportar_Click()

End Sub

' Botón de Listar
Private Sub cmdListar_Click()

End Sub

' Botón de Limpiar
Private Sub cmdLimpiar_Click()

End Sub

' Al hacer click sobre el flex de datos
Private Sub flxDatos_Click()
  Dim i As Integer
  For i = 0 To flxDatos.Cols - 1
    flxEdicionCampos.col = 1
    flxEdicionCampos.Row = i + 1
    flxEdicionCampos.Text = flxDatos.TextMatrix(flxDatos.Row, i)
  Next
End Sub

' Para poder utilizar el Flexgrid como una caja de texto
Private Sub flxEdicionCampos_KeyPress(KeyAscii As Integer)
' Si presiona Espacio en blanco
  If KeyAscii = 32 Then
    flxEdicionCampos.Text = flxEdicionCampos.Text & Chr(KeyAscii)
    miCadena = miCadena & Chr(KeyAscii)
  End If
  ' Si presiona Letras Mayusculas
  If KeyAscii >= 65 And KeyAscii <= 90 Then
    flxEdicionCampos.Text = flxEdicionCampos.Text & Chr(KeyAscii)
    miCadena = miCadena & Chr(KeyAscii)
  End If
  ' Si presiona Letras Minusculas
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    flxEdicionCampos.Text = flxEdicionCampos.Text & Chr(KeyAscii)
    miCadena = miCadena & Chr(KeyAscii)
  End If
  ' Si presiona Números
  If KeyAscii > 47 And KeyAscii < 58 Then
    flxEdicionCampos.Text = flxEdicionCampos.Text & Chr(KeyAscii)
    miCadena = miCadena & Chr(KeyAscii)
  End If
  ' Cuando se presiona Enter
  If KeyAscii = 13 Then  'Presiona ENTER
    'Aca pones la rutina para cuando presiona ENTER
    KeyAscii = 0  'Esto elimina el Beep del Enter
  End If
End Sub



'    rst.Open "select " & _
     '            "T3.Nombre as Nombre_y_Apellido,  " & _
     '            "T3.NroCarnet as Afiliado, " & _
     '            "(((365 * year(getdate()))-(365*(year(T3.FecNac))))+ (month(getdate())-month(T3.FecNac))*30 + (day(getdate()) - day(T3.FecNac)))/365 as Edad, " & _
     '            "T2.FecDes as Ingreso, " & _
     '            "T1.NroCama as Numero_Cama, " & _
     '            "T1.LetCama as Letra_Cama, " & _
     '            "T1.Area as Area, " & _
     '            "T2.DiagIngDes as Diagnóstico, " & _
     '            "T2.IDCli as Cliente " & _
     '            "from Camas as T1 " & _
     '            "inner join CabMovInter as T2 on T1.IDCuenta = T2.IDCuenta " & _
     '            "inner join Padron as T3 on T2.IDPad = T3.IDPad " & _
     '            "Where T1.IDCuenta <> 0 And T2.IDCli = 509 ", oConexion.cnn, adOpenDynamic, adLockOptimistic

'    rst.Open "select p.nombre, p.numdoc, c.idcuenta, c.fecdes, c.fecHas, pre.idpre, n.abrev, pre.cant, pre.fecefe, " & _
     '            "dbo.ObtenerCliente_PorIDPlan (pre.idplan) from prestacionesinter pre " & _
     '            "inner join cabmovinter c " & _
     '            "on c.idcuenta = pre.idcuenta " & _
     '            "inner join padron p " & _
     '            "on p.idpad = c.idpad " & _
     '            "inner join nomenclador n " & _
     '            "on n.idpre = pre.idpre " & _
     '            "where pre.act = 1 and pre.idliq > 0 and " & _
     '            "pre.sefact = 'S' and " & _
     '            "convert(char(8),pre.fecefe,112) between '20190101' and '20190731' " & _
     '            "order by p.nombre asc, pre.fecefe asc", oConexion.cnn, adOpenDynamic, adLockOptimistic

'   rst.Open "select " & _
    '            "pre.idpre, n.abrev, count(*) " & _
    '            "from prestacionesinter pre " & _
    '            "inner join nomenclador n " & _
    '            "on n.idpre = pre.idpre " & _
    '            "where pre.act = 1 and pre.idliq > 0 and " & _
    '            "pre.sefact = 'S' and " & _
    '            "pre.idpre > '500000' and pre.idpre < '570000' and " & _
    '            "convert(char(8),pre.fecefe,112) between '20190101' and '20190731' " & _
    '            "group by pre.idpre, n.abrev " & _
    '            "order by pre.idpre asc", oConexion.cnn, adOpenDynamic, adLockOptimistic


''''''
''''''  If Not (rst.BOF And rst.EOF) Then
''''''  While Not (rst.EOF) And miFil < Val(txtRegistros) + 2
''''''    flxDatos.Rows = miFil + 1
''''''    'flxDatos.Row = 1
''''''    col = 0
''''''    For Each tField In rst.Fields
''''''      flxDatos.col = col
''''''      If IsNull(rst(col)) Then
''''''        flxDatos.Text = ""
''''''      Else
''''''        If IsNumeric(rst(col)) Then
''''''          flxDatos.Text = Str(rst(col))
''''''        Else
''''''          If rst(col) = "U" Then
''''''            flxDatos.Text = "Terapia"
''''''          ElseIf rst(col) = "P" Then
''''''            flxDatos.Text = "Piso"
''''''          Else
''''''            flxDatos.Text = rst(col)
''''''          End If
''''''        End If
''''''      End If
''''''      col = col + 1
''''''    Next
''''''    rst.MoveNext
''''''    miFil = miFil + 1
''''''    flxDatos.Row = flxDatos.Row + 1
''''''  Wend
''''''
''''''  ' Quita la ultima fila
''''''  flxDatos.RemoveItem (flxDatos.Row)
''''''  End If
''''''  flxDatos.Refresh
''''''  flxDatos.Visible = True
''''''  lblAviso.Visible = False
''''''  DoEvents
