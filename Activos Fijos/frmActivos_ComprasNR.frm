VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmActivos_ComprasNR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Compras de Activos Fijos No Registrados"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   14415
      _Version        =   1441793
      _ExtentX        =   25426
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   14415
      _Version        =   1441793
      _ExtentX        =   25426
      _ExtentY        =   11456
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Left            =   8280
      Top             =   480
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption lblY 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14655
      _Version        =   1441793
      _ExtentX        =   25850
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   $"frmActivos_ComprasNR.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Periodo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblX 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10398
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmActivos_ComprasNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtpFecha_Change()
 
lblX.Caption = fxActivos_Periodo(dtpFecha.Value)
Call TimerX_Timer
 
End Sub

Private Sub Form_Load()
vModulo = 36

With lsw.ColumnHeaders

    .Clear
    .Add , , "No. Factura", 3200
    .Add , , "Línea", 1200, vbCenter
    .Add , , "Proveedor", 3200
    .Add , , "Activo Nombre", 3200
    .Add , , "Cantidad", 1200, vbCenter
    .Add , , "Registrada", 1200, vbCenter
    .Add , , "Pendiente", 1200, vbCenter
    .Add , , "Valor Adq.", 2200, vbRightJustify

End With


dtpFecha.Value = gActivos.Periodo
Call dtpFecha_Change

End Sub


Private Sub lsw_DblClick()

If lsw.ListItems.Count <= 0 Then Exit Sub

On Error GoTo vError

gAsistente.Documento = lsw.SelectedItem.Text
gAsistente.Proveedor = lsw.SelectedItem.Tag
gAsistente.VU = CCur(lsw.SelectedItem.SubItems(7))
gAsistente.Tipo = "C"

Call sbFormsCall("frmActivos_Main", , , , , Me, True)

vError:


End Sub


Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaI As Date, vFechaC As Date, vRegistro As Currency
Dim itmX As ListViewItem


TimerX.Interval = 0

vFechaI = CDate(Year(dtpFecha.Value) & "/" & Format(Month(dtpFecha.Value), "00") & "/01")
vFechaC = DateAdd("d", -1, DateAdd("m", 1, vFechaI))

strSQL = "SELECT D.COD_FACTURA,D.LINEA,D.COD_PROVEEDOR,D.COD_PRODUCTO,(D.PRECIO * ((D.IMP_VENTAS / 100) + 1)) AS PRECIO" _
       & ",D.CANTIDAD,P.DESCRIPCION AS PROVEEDOR,P.DESCRIPCION AS PRODUCTO,E.FECHA" _
       & " FROM CPR_COMPRAS E inner join CPR_COMPRAS_detalle D" _
       & " on E.cod_factura = D.cod_factura and E.cod_proveedor = D.cod_proveedor" _
       & " inner join pv_productos P on D.cod_producto = P.cod_producto" _
       & " and P.tipo_producto = 'A'" _
       & " inner join cxp_proveedores X on D.cod_proveedor = X.cod_proveedor" _
       & " WHERE E.FECHA BETWEEN '" & Format(vFechaI, "yyyy/mm/dd") & "' AND '" _
       & Format(vFechaC, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear
PrgBar.Max = rs.RecordCount + 2
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF

 vRegistro = fxActivos_RegistroCompras(vFechaI, vFechaC, rs!COD_PROVEEDOR, rs!cod_factura)
 
 If rs!cantidad > vRegistro Then
    Set itmX = lsw.ListItems.Add(, , rs!cod_factura)
        itmX.Tag = rs!COD_PROVEEDOR
        itmX.SubItems(1) = rs!Linea
        itmX.SubItems(2) = rs!Proveedor
        itmX.SubItems(3) = rs!producto
        itmX.SubItems(4) = rs!cantidad
        itmX.SubItems(5) = vRegistro
        itmX.SubItems(6) = rs!cantidad - vRegistro
        itmX.SubItems(7) = rs!Precio
   
 End If
 PrgBar.Value = PrgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

PrgBar.Visible = False

End Sub
