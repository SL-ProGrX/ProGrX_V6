VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmInvRepGeneral 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inv: Reportes Generales"
   ClientHeight    =   7392
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8100
   Icon            =   "frmInvRepGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7392
   ScaleWidth      =   8100
   Begin VB.ComboBox cboProveedor 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   5760
      Width           =   5415
   End
   Begin VB.ComboBox cboUnidad 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ComboBox cboDepartamento 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   5400
      Width           =   5415
   End
   Begin VB.ComboBox cboLinea 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   5040
      Width           =   5415
   End
   Begin VB.ComboBox cboBodega 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   4680
      Width           =   5415
   End
   Begin VB.ComboBox cboProExistencia 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ComboBox cboProEstado 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ComboBox cboProTipo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   1812
      Left            =   360
      TabIndex        =   16
      Top             =   1320
      Width           =   7332
      _ExtentX        =   12933
      _ExtentY        =   3196
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reporte"
         Object.Width           =   8008
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   360
      TabIndex        =   20
      Top             =   6240
      Width           =   7452
      _Version        =   1245187
      _ExtentX        =   13144
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin VB.ComboBox cboOrden 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   276
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   5280
         TabIndex        =   21
         Top             =   360
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   14
         Picture         =   "frmInvRepGeneral.frx":030A
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ordernar por"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   732
      Left            =   2160
      TabIndex        =   24
      Top             =   240
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Informes Generales de Inventarios "
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   252
      Left            =   360
      TabIndex        =   19
      Top             =   3240
      Width           =   7332
      _Version        =   1245187
      _ExtentX        =   12933
      _ExtentY        =   444
      _StockProps     =   14
      Caption         =   "Parámetros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Reporte >>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   360
      TabIndex        =   18
      Top             =   3600
      Width           =   1452
   End
   Begin VB.Label lblReporte 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   1920
      TabIndex        =   17
      Top             =   3600
      Width           =   5412
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   360
      TabIndex        =   14
      Top             =   5760
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidades"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   3960
      TabIndex        =   9
      Top             =   4320
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Departamentos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   5400
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   4680
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Existencia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   3960
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado del Producto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Producto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmInvRepGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLlenaCbos()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

'Carga Bodegas
Call sbPosCombosCarga("Bodegas", cboBodega)
cboBodega.AddItem "[TODOS]"
cboBodega.Text = "[TODOS]"


'Carga Lineas
cboLinea.Clear
strSQL = "select cod_prodclas,descripcion from PV_PROD_CLASIFICA"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 cboLinea.AddItem Trim(rs!Descripcion)
 cboLinea.ItemData(cboLinea.NewIndex) = rs!cod_prodclas
 rs.MoveNext
Loop
rs.Close
cboLinea.AddItem "[TODOS]"
cboLinea.Text = "[TODOS]"

'Carga Proveedores
cboProveedor.Clear
strSQL = "select cod_proveedor,descripcion from CXP_Proveedores"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 cboProveedor.AddItem Trim(rs!Descripcion)
 cboProveedor.ItemData(cboProveedor.NewIndex) = rs!cod_proveedor
 rs.MoveNext
Loop
rs.Close
cboProveedor.AddItem "[TODOS]"
cboProveedor.Text = "[TODOS]"


'Carga Unidades
cboUnidad.Clear
strSQL = "select cod_unidad,descripcion from PV_UNIDADES"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 cboUnidad.AddItem Trim(rs!cod_unidad) & " - " & Trim(rs!Descripcion)
 rs.MoveNext
Loop
rs.Close
cboUnidad.AddItem "[TODOS]"
cboUnidad.Text = "[TODOS]"

'Carga Departamentos
cboDepartamento.Clear
strSQL = "select cod_departamento,descripcion from PV_DEPARTAMENTOS"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 cboDepartamento.AddItem Trim(rs!cod_departamento) & " - " & Trim(rs!Descripcion)
 rs.MoveNext
Loop
rs.Close
cboDepartamento.AddItem "[TODOS]"
cboDepartamento.Text = "[TODOS]"

'Carga Estados del Producto
cboProEstado.Clear
cboProEstado.AddItem "Activos"
cboProEstado.AddItem "InActivos"
cboProEstado.AddItem "[TODOS]"
cboProEstado.Text = "[TODOS]"

'Carga Existencias Rangos
Call sbInvExistenciaCargaCbo(cboProExistencia)

'Carga Tipos de Productos
cboProTipo.Clear
cboProTipo.AddItem "Producto"
cboProTipo.AddItem "Servicio"
cboProTipo.AddItem "[TODOS]"
cboProTipo.Text = "[TODOS]"

'Ordenamiento
cboOrden.Clear
cboOrden.AddItem "01 - Descripción"
cboOrden.AddItem "02 - Código"
cboOrden.AddItem "03 - Barras [Productos]"
cboOrden.Text = "02 - Código"

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub sbAICbos()
cboProEstado.Enabled = True
cboProExistencia.Enabled = True
cboProTipo.Enabled = True

cboBodega.Enabled = False
cboLinea.Enabled = False
cboDepartamento.Enabled = False
cboUnidad.Enabled = False
cboProveedor.Enabled = False

cboOrden.Visible = False
lbl.Visible = cboOrden.Visible

Select Case lblReporte.Tag
  Case "y00" 'Inventario General
    cboLinea.Enabled = True
    cboLinea.Text = "[TODOS]"
    cboUnidad.Enabled = True
    cboUnidad.Text = "[TODOS]"
  Case "y01" 'Inventario por Bodega
    cboBodega.Enabled = True
    cboBodega.Text = "[TODOS]"
    cboLinea.Enabled = True
    cboLinea.Text = "[TODOS]"
    cboUnidad.Enabled = True
    cboUnidad.Text = "[TODOS]"
  Case "y02" 'Productos por Líneas
    cboLinea.Enabled = True
    cboLinea.Text = "[TODOS]"
  Case "y03" 'Productos por Departamento
    cboDepartamento.Enabled = True
    cboDepartamento.Text = "[TODOS]"
  Case "y04" 'Productos por Unidades
    cboUnidad.Enabled = True
    cboUnidad.Text = "[TODOS]"
  Case "y05" 'Productos por Proveedor
    cboProveedor.Enabled = True
    cboProveedor.Text = "[TODOS]"
  Case "y06" 'Líneas por Departamento
    cboDepartamento.Enabled = True
    cboDepartamento.Text = "[TODOS]"
  Case "y07" 'Departamentos por Lineas
    cboLinea.Enabled = True
    cboLinea.Text = "[TODOS]"
End Select

End Sub



Private Sub btnReporte_Click()
Call sbReporte
End Sub

Private Sub cboBodega_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Convertir = ""
gBusquedas.Columna = "descripcion"
gBusquedas.Orden = "descripcion"
gBusquedas.Consulta = "Select cod_bodega as Codigo,descripcion from pv_bodegas"
gBusquedas.Filtro = ""
frmBusquedas.Show vbModal

cboBodega.Text = Trim(gBusquedas.Resultado) & " - " & Trim(gBusquedas.Resultado2)
Exit Sub
vError:
  cboBodega.Text = "[TODOS]"

End Sub


Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Convertir = ""
gBusquedas.Columna = "descripcion"
gBusquedas.Orden = "descripcion"
gBusquedas.Consulta = "Select cod_departamento as Codigo,descripcion from pv_departamentos"
gBusquedas.Filtro = ""
frmBusquedas.Show vbModal

cboDepartamento.Text = Trim(gBusquedas.Resultado) & " - " & Trim(gBusquedas.Resultado2)
Exit Sub
vError:
  cboDepartamento.Text = "[TODOS]"

End Sub

Private Sub cboLinea_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Convertir = ""
gBusquedas.Columna = "descripcion"
gBusquedas.Orden = "descripcion"
gBusquedas.Consulta = "Select Cod_ProdClas as Codigo,descripcion from pv_prod_clasifica"
gBusquedas.Filtro = ""
frmBusquedas.Show vbModal

cboLinea.Text = Trim(gBusquedas.Resultado2)
Exit Sub
vError:
  cboLinea.Text = "[TODOS]"
End Sub


Private Sub cboProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Convertir = ""
gBusquedas.Columna = "descripcion"
gBusquedas.Orden = "descripcion"
gBusquedas.Consulta = "Select cod_proveedor as Codigo,descripcion from cxp_proveedores"
gBusquedas.Filtro = ""
frmBusquedas.Show vbModal

cboProveedor.Text = Trim(gBusquedas.Resultado2)
Exit Sub

vError:
  cboProveedor.Text = "[TODOS]"

End Sub

Private Sub cboUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Convertir = ""
gBusquedas.Columna = "descripcion"
gBusquedas.Orden = "descripcion"
gBusquedas.Consulta = "Select cod_unidad as Codigo,descripcion from pv_unidades"
gBusquedas.Filtro = ""
frmBusquedas.Show vbModal

cboUnidad.Text = Trim(gBusquedas.Resultado) & " - " & Trim(gBusquedas.Resultado2)
Exit Sub
vError:
  cboUnidad.Text = "[TODOS]"

End Sub

Private Sub Form_Load()


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


lsw.ListItems.Clear
lsw.ListItems.Add , "x00", "Listado de Productos"
lsw.ListItems.Add , "x01", "Listado de Bodegas"
lsw.ListItems.Add , "x02", "Listado de Departamentos"
lsw.ListItems.Add , "x03", "Listado de Líneas de Productos"
lsw.ListItems.Add , "x04", "Listado de Unidades de medida"
lsw.ListItems.Add , "x05", "Listado de Tipos de Precios"
lsw.ListItems.Add , "x06", "Listado de Transacciones E/S/T"
lsw.ListItems.Add , "y00", "Inventario General"
lsw.ListItems.Add , "y01", "Inventario x Bodegas"
lsw.ListItems.Add , "y02", "Productos x líneas"
lsw.ListItems.Add , "y03", "Productos x Departamentos"
lsw.ListItems.Add , "y04", "Productos x Unidades"
lsw.ListItems.Add , "y05", "Productos x Proveedor"
lsw.ListItems.Add , "y06", "Líneas x Departamentos"
lsw.ListItems.Add , "y07", "Departamentos x Líneas"

lblReporte.Tag = "x00"
lblReporte.Caption = "Listado de Productos"


Call sbLlenaCbos

Call sbAICbos
cboProEstado.Enabled = False
cboProExistencia.Enabled = False
cboProTipo.Enabled = False

cboOrden.Visible = True
lbl.Visible = True

End Sub

Private Sub lsw_Click()
lblReporte.Tag = lsw.SelectedItem.Key
lblReporte.Caption = lsw.SelectedItem
Call sbAICbos

If Mid(lblReporte.Tag, 1, 1) = "x" Then
    cboProEstado.Enabled = False
    cboProExistencia.Enabled = False
    cboProTipo.Enabled = False
    cboOrden.Visible = True
    lbl.Visible = cboOrden.Visible
End If

End Sub


Private Function fxSQL(Optional i As Integer = 0) As String
Dim vSQL As String

vSQL = ""

'En todos los reportes se exponen los parametros siguientes
If cboProTipo.Enabled And cboProTipo.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_PRODUCTOS.TIPO_PRODUCTO} = '" & Mid(cboProTipo.Text, 1, 1) & "'"
End If

If cboProEstado.Enabled And cboProEstado.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_PRODUCTOS.ESTADO} = '" & Mid(cboProEstado.Text, 1, 1) & "'"
End If

If cboProExistencia.Enabled And cboProExistencia.Text <> "[TODOS]" Then
  If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
  Select Case i
    Case 0 'productos
      Select Case Mid(cboProExistencia.Text, 1, 2)
        Case "00" 'Agotados
           vSQL = vSQL & " {PV_PRODUCTOS.EXISTENCIA} = 0"
        Case "01" 'Minima
           vSQL = vSQL & " {PV_PRODUCTOS.EXISTENCIA} <= {PV_PRODUCTOS.INVENTARIO_MINIMO}"
        Case "02" 'Maxima
           vSQL = vSQL & " {PV_PRODUCTOS.EXISTENCIA} >= {PV_PRODUCTOS.INVENTARIO_MAXIMO}"
        Case "03" 'Inv (-) Reposición
           vSQL = vSQL & " {PV_PRODUCTOS.EXISTENCIA} < 0"
        Case "04" 'Mayor Igual xx
           vSQL = vSQL & " {PV_PRODUCTOS.EXISTENCIA} >= "
           vSQL = vSQL & InputBox("Existencia Mayor / Igual a ?", "Reportes de Inventarios")

      End Select
    
    Case 1 'Bodegas
    
      Select Case Mid(cboProExistencia.Text, 1, 2)
        Case "00" 'Agotados
           vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} = 0"
        Case "01" 'Minima
           vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} <= {PV_PRODUCTOS.INVENTARIO_MINIMO}"
        Case "02" 'Maxima
           vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} >= {PV_PRODUCTOS.INVENTARIO_MAXIMO}"
        Case "03" 'Inv (-) Reposición
           vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} < 0"
        Case "04" 'Mayor Igual xx
           vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} >= "
           vSQL = vSQL & InputBox("Existencia Mayor / Igual a ?", "Reportes de Inventarios")
      End Select
  End Select
      
End If

'*****

If cboLinea.Enabled And cboLinea.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   Select Case i
     Case 0, 1
       vSQL = vSQL & " {PV_PRODUCTOS.COD_PRODCLAS} = " & cboLinea.ItemData(cboLinea.ListIndex)
     Case 3
       vSQL = vSQL & " {PV_PROD_CLASIFICA.COD_PRODCLAS} = " & cboLinea.ItemData(cboLinea.ListIndex)
   End Select
End If

If cboUnidad.Enabled And cboUnidad.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_PRODUCTOS.COD_UNIDAD} = '" & fxCodigoCbo(cboUnidad) & "'"
End If
     
If cboBodega.Enabled And cboBodega.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_BODEGAS.COD_BODEGA} = '" & fxCodigoCbo(cboBodega) & "'"
End If

If cboDepartamento.Enabled And cboDepartamento.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_DEPARTAMENTOS.COD_DEPARTAMENTO} = '" & fxCodigoCbo(cboDepartamento) & "'"
End If

If cboProveedor.Enabled And cboProveedor.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {CXP_PROVEEDORES.COD_PROVEEDOR} = " & cboProveedor.ItemData(cboProveedor.ListIndex)
End If



fxSQL = vSQL

End Function


Private Sub sbReporte()
Dim vSQL As String

vSQL = ""


Select Case lblReporte.Tag
  
  Case "x00" 'Listado de Productos
       Call sbInvReportes("ProductosGen", lblReporte.Caption, "Productos/Acticulos y Servicios", vSQL)
  Case "x01" 'Listado de Bodegas
     Call sbInvReportes("Bodegas", lblReporte.Caption, "Listado", "")
  Case "x02" 'Listado de Departamentos
     Call sbInvReportes("Departamentos", lblReporte.Caption, "Listado", "")
  Case "x03" 'Listado de Lineas de Prod
     Call sbInvReportes("TiposProductos", lblReporte.Caption, "Listado", "")
  Case "x04" 'Listado de Unidades
     Call sbInvReportes("Unidades", lblReporte.Caption, "Listado", "")
  Case "x05" 'Listado de Tipos de Precios
     Call sbInvReportes("TiposPrecios", lblReporte.Caption, "Listado", "")
  Case "x06" 'Listado de E/S/T
     Call sbInvReportes("TiposEST", lblReporte.Caption, "Listado", "")
  
  Case "y00" 'Inventario General
     Call sbInvReportes("ProductosInv", lblReporte.Caption, "Productos/Acticulos y Servicios", fxSQL(0))
  Case "y01" 'Inventario x Bodegas
     Call sbInvReportes("InvBodegas", lblReporte.Caption, "Inventario", fxSQL(1))
  Case "y02" 'Productos x Lineas
     Call sbInvReportes("ProductosLinea", lblReporte.Caption, "Listado", fxSQL)
  Case "y03" 'Productos x Departamentos
     Call sbInvReportes("ProductosDept", lblReporte.Caption, "Listado", fxSQL)
  Case "y04" 'Productos x Unidades
     Call sbInvReportes("ProductosUnidades", lblReporte.Caption, "Listado", fxSQL)
  Case "y05" 'Productos x Proveedor
     Call sbInvReportes("ProductosProveedor", lblReporte.Caption, "Listado de Proveedores vrs Productos", fxSQL)
  Case "y06" 'Lineas x Departamentos
     Call sbInvReportes("DeptLineas", lblReporte.Caption, "Listado", fxSQL)
  Case "y07" 'Departamentos x Lineas
     Call sbInvReportes("LineasDept", lblReporte.Caption, "Listado", fxSQL(3))

End Select

End Sub
