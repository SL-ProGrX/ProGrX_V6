VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmInvKardex 
   Caption         =   "Kardex"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkProductos 
      Height          =   252
      Left            =   7560
      TabIndex        =   12
      Top             =   120
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4092
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   12012
      _Version        =   524288
      _ExtentX        =   21188
      _ExtentY        =   7218
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   485
      SpreadDesigner  =   "frmInvKardex.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   612
      Left            =   6240
      TabIndex        =   8
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmInvKardex.frx":0A8F
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   7560
      TabIndex        =   9
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmInvKardex.frx":14AD
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1080
      TabIndex        =   10
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   3240
      TabIndex        =   11
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkBodegas 
      Height          =   252
      Left            =   7560
      TabIndex        =   13
      Top             =   480
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   1080
      TabIndex        =   14
      Top             =   1320
      Width           =   3492
      _Version        =   1441793
      _ExtentX        =   6165
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16579836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16579836
      Style           =   2
      Appearance      =   14
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboRep 
      Height          =   312
      Left            =   1080
      TabIndex        =   15
      Top             =   1680
      Width           =   3492
      _Version        =   1441793
      _ExtentX        =   6165
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16579836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16579836
      Style           =   2
      Appearance      =   14
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnMovimientos 
      Height          =   315
      Left            =   4680
      TabIndex        =   16
      Top             =   1680
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Mov.Bodegas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1080
      TabIndex        =   17
      Top             =   120
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   312
      Left            =   4680
      TabIndex        =   18
      Top             =   1320
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "1000"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3000
      TabIndex        =   19
      Top             =   120
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7853
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox cboBodega 
      Height          =   312
      Left            =   1080
      TabIndex        =   20
      Top             =   480
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16579836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16579836
      Style           =   2
      Appearance      =   14
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "# Líneas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bodegas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "frmInvKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSubTitulo As String

Private Sub sbInicializa()
Dim strSQL As String

txtCodigo = ""
txtNombre = ""

Call sbPosCombosCarga("bodegas", cboBodega)

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Call sbInvOrigenCargaCbo(cboTipo)

cboRep.Clear
cboRep.AddItem "01 - Agr. x Bodegas / Tipo"
cboRep.AddItem "02 - Agr. x Bodegas / Origen"
cboRep.AddItem "03 - Un Producto x Bodegas / Tipo"
cboRep.AddItem "04 - Un Producto x Bodegas / Origen"
cboRep.AddItem "05 - Un Producto x Bodegas / General"

cboRep.Text = "01 - Agr. x Bodegas / Tipo"


Call chkProductos_Click
Call chkBodegas_Click
End Sub

Private Sub btnMovimientos_Click()
frmInvExistenciaProducto.Show
End Sub

Private Sub chkBodegas_Click()
If chkBodegas.Value = vbChecked Then
  cboBodega.Enabled = False
Else
  cboBodega.Enabled = True
End If
End Sub

Private Sub chkProductos_Click()
If chkProductos.Value = vbChecked Then
   txtCodigo.Enabled = False
   txtNombre.Enabled = False
Else
   txtCodigo.Enabled = True
   txtNombre.Enabled = True
End If
End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Not IsNumeric(txtLineas) Then txtLineas = 100

strSQL = "select Top " & Trim(txtLineas) & " M.Fecha,(rtrim(M.cod_producto) + ' - ' + rtrim(P.descripcion)) as Producto" _
       & ",case M.tipo when 'E' then 'ENTRADA' when 'S' then 'SALIDA' end as TipoX" _
       & ",M.origen,M.codigo,isnull(M.existencia,0) as Existencia,M.cantidad" _
       & ",case when M.tipo = 'E' then isnull(M.existencia,0) + M.Cantidad when M.tipo = 'S' then isnull(M.existencia,0) - M.Cantidad end as ExistenciaX" _
       & ",M.precio,(M.cantidad * M.precio) as TotalSinImp,(M.cantidad * M.precio) * (M.imp_ventas / 100) as ImpVentas" _
       & ",(M.cantidad * M.precio) * (M.imp_consumo / 100) as ImpConsumo" _
       & ",(M.cantidad * M.precio) + ((M.cantidad * M.precio) * (M.imp_ventas / 100)) + ((M.cantidad * M.precio) * (M.imp_consumo / 100)) as TotalConImp" _
       & ",(rtrim(M.cod_bodega) + ' - ' + rtrim(B.descripcion)) as Bodega" _
       & ",dbo.fxINVBodegaTraslado(M.Origen,M.Tipo,M.Linea) as BodegaEnlace" _
       & " from pv_inventario_mov M inner join pv_productos P on M.cod_producto = P.cod_producto" _
       & " inner join pv_Bodegas B on M.cod_bodega = B.cod_bodega" _
       & " where M.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
       
If chkProductos.Value = vbUnchecked Then
  strSQL = strSQL & " and M.cod_producto = '" & txtCodigo & "'"
End If

If chkBodegas.Value = vbUnchecked Then
  strSQL = strSQL & " and M.cod_bodega = '" & cboBodega.ItemData(cboBodega.ListIndex) & "'"
End If

Select Case cboTipo.Text
 Case "[SOLO SALIDAS]"
   strSQL = strSQL & " and M.tipo = 'S'"
 Case "[SOLO ENTRADAS]"
   strSQL = strSQL & " and M.tipo = 'E'"
 Case "[TODOS]"
   'Nada
 Case Else
   strSQL = strSQL & " and M.origen = '" & cboTipo.Text & "'"
End Select
      
strSQL = strSQL & " order by M.Fecha desc"
      
Call sbCargaGrid(vGrid, 15, strSQL)
      
vGrid.MaxRows = vGrid.MaxRows - 1
      
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Function fxSQL() As String
Dim vRes As String

vRes = "{PV_INVENTARIO_MOV.FECHA} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
     & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

vSubTitulo = "[ INICIO : " & Format(dtpInicio.Value, "dd/mm/yyyy") & " CORTE : " & Format(dtpCorte.Value, "dd/mm/yyyy") & " ] "

If chkProductos.Value = vbUnchecked Then
  vRes = vRes & " AND {PV_INVENTARIO_MOV.COD_PRODUCTO} = '" & txtCodigo & "'"
  vSubTitulo = vSubTitulo & "[ PRODUCTO : " & Trim(txtCodigo) & " - " & Trim(txtNombre) & " ] "
End If

If chkBodegas.Value = vbUnchecked Then
  vRes = vRes & " AND {PV_INVENTARIO_MOV.COD_BODEGA} = '" & cboBodega.ItemData(cboBodega.ListIndex) & "'"
  vSubTitulo = vSubTitulo & "[ BODEGA : " & cboBodega.Text & " ] "
End If

Select Case cboTipo.Text
 Case "[SOLO SALIDAS]"
   vRes = vRes & " AND {PV_INVENTARIO_MOV.TIPO} = 'S'"
 Case "[SOLO ENTRADAS]"
   vRes = vRes & " AND {PV_INVENTARIO_MOV.TIPO} = 'E'"
 Case "[TODOS]"
   'Nada
 Case Else
   vRes = vRes & " AND {PV_INVENTARIO_MOV.ORIGEN} = '" & cboTipo.Text & "'"
End Select

vSubTitulo = vSubTitulo & "[ TIPO : " & UCase(cboTipo.Text) & " ]"

fxSQL = vRes

End Function

Private Sub cmdReporte_Click()
Dim strSQL As String

If chkProductos.Value = vbChecked And Val(SIFGlobal.fxCodText(cboRep.Text)) > 2 Then
  MsgBox "Este tipo de reporte no aplica para multiples Productos/Articulos o Servicios", vbExclamation
  Exit Sub
End If


Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de Inventarios ¦ Movimientos (kardex)"
 
 .Connect = glogon.ConectRPT
 
 strSQL = fxSQL
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxTitulo = 'KARDEX'"
 .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Inventario_KardexN" & SIFGlobal.fxCodText(cboRep.Text) & ".rpt")
 
 .SelectionFormula = strSQL
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

vGrid.AppearanceStyle = fxGridStyle

Call sbInicializa

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 350
vGrid.Height = Me.Height - 2690

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  frmBusquedaArticulos.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCodigo_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCodigo, "Productos")
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkProductos.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_producto,descripcion from pv_productos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub


