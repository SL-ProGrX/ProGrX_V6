VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCajas_ConsultaMovimientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cajas: Consulta de Movimientos por Forma de Pago"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   14655
      _Version        =   524288
      _ExtentX        =   25850
      _ExtentY        =   10186
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
      MaxCols         =   18
      SpreadDesigner  =   "frmCajas_ConsultaMovimientos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpRegistroInicio 
      Height          =   312
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   1332
      _Version        =   1441792
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
   Begin XtremeSuiteControls.DateTimePicker dtpRegistroCorte 
      Height          =   312
      Left            =   2400
      TabIndex        =   9
      Top             =   1680
      Width           =   1332
      _Version        =   1441792
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   2652
      _Version        =   1441792
      _ExtentX        =   4678
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   5160
      TabIndex        =   11
      Top             =   1320
      Width           =   2052
      _Version        =   1441792
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   5160
      TabIndex        =   12
      Top             =   1680
      Width           =   2052
      _Version        =   1441792
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNumDoc 
      Height          =   315
      Left            =   8280
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
      _Version        =   1441792
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboFormaPago 
      Height          =   312
      Left            =   8280
      TabIndex        =   14
      Top             =   1320
      Width           =   2292
      _Version        =   1441792
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkSF_Relacionados 
      Height          =   252
      Left            =   6720
      TabIndex        =   15
      Top             =   2040
      Width           =   3852
      _Version        =   1441792
      _ExtentX        =   6794
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar los saldos a favor relacionados?  "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   252
      Left            =   1320
      TabIndex        =   16
      Top             =   2040
      Width           =   2412
      _Version        =   1441792
      _ExtentX        =   4254
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas las Fechas?  "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   420
      Left            =   10680
      TabIndex        =   17
      Top             =   1320
      Width           =   1350
      _Version        =   1441792
      _ExtentX        =   2381
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Buscar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_ConsultaMovimientos.frx":0B73
   End
   Begin XtremeSuiteControls.ComboBox cboCajas 
      Height          =   312
      Left            =   1080
      TabIndex        =   18
      Top             =   2400
      Width           =   2652
      _Version        =   1441792
      _ExtentX        =   4683
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCajaAP 
      Height          =   312
      Left            =   5160
      TabIndex        =   20
      Top             =   2400
      Width           =   2052
      _Version        =   1441792
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   420
      Left            =   12000
      TabIndex        =   22
      Top             =   1320
      Width           =   1350
      _Version        =   1441792
      _ExtentX        =   2381
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Exportar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_ConsultaMovimientos.frx":1273
   End
   Begin XtremeSuiteControls.ComboBox cboMov 
      Height          =   315
      Left            =   8280
      TabIndex        =   23
      Top             =   1680
      Width           =   2295
      _Version        =   1441792
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   315
      Left            =   11040
      TabIndex        =   25
      Top             =   2400
      Width           =   2295
      _Version        =   1441792
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   11040
      TabIndex        =   26
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo Mov.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7320
      TabIndex        =   24
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Id Apertura.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   3960
      TabIndex        =   21
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Caja .:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No. Doc.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "F. Pago.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   7320
      TabIndex        =   6
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nombre.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Id / Cédula.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha .:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Usuario.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Afectaciones (Movimientos) por Forma de Pago"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   9585
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "frmCajas_ConsultaMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean



Private Sub btnBuscar_Click()
     Call sbConsulta
End Sub

Private Function fxCajasUltimaApertura(pCajas As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim Resultado As Long

On Error GoTo vError

Resultado = 0

strSQL = "select dbo.fxSIFDocsCajaUltimaApertura('" & pCajas & "') as Resultado"
Call OpenRecordSet(rs, strSQL)
    Resultado = rs!Resultado
rs.Close

vError:

fxCajasUltimaApertura = Resultado

End Function


Private Sub btnExport_Click()
 Dim vHeaders As vGridHeaders
 
    vHeaders.Columnas = vGrid.MaxCols
    

    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Tipo Doc."
    vHeaders.Headers(4) = "Num. Doc."
    vHeaders.Headers(5) = "Monto Aplicado"
    vHeaders.Headers(6) = "Divisa"
    vHeaders.Headers(7) = "Tipo Cambio"
    vHeaders.Headers(8) = "Reg.Fecha"
    vHeaders.Headers(9) = "Reg.Usuario"
    vHeaders.Headers(10) = "Forma Pago"
    vHeaders.Headers(11) = "Referencia"
    vHeaders.Headers(12) = "Banco"
    vHeaders.Headers(13) = "Pagador"
    vHeaders.Headers(14) = "Cuenta"
    vHeaders.Headers(15) = "Concepto"
    vHeaders.Headers(16) = "Caja"
    vHeaders.Headers(17) = "Apertura/Cierre"
    
    
    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Cajas_Consulta_Movimientos")


End Sub

Private Sub cboCajas_Click()

If vPaso Then Exit Sub
If cboCajas.ListCount = 0 Then Exit Sub

If cboCajas.Text = "TODOS" Then
    txtCajaAP.Text = ""
Else
    txtCajaAP.Text = fxCajasUltimaApertura(cboCajas.ItemData(cboCajas.ListIndex))
End If

End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpRegistroInicio.Enabled = False
Else
   dtpRegistroInicio.Enabled = True
End If

dtpRegistroCorte.Enabled = dtpRegistroInicio.Enabled
  
End Sub

Private Sub Form_Activate()
 vModulo = 5

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 5

'Carga las cuentas bancarias asiganadas a la forma de pago
vPaso = True

Me.Width = 12630
Me.Height = 9465

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

strSQL = "select  rtrim(COD_FORMA_PAGO) as 'IdX', rtrim(DESCRIPCION) as 'itmX' from SIF_FORMAS_PAGO" _
       & " WHERE ACTIVA = 1 ORDER BY COD_FORMA_PAGO"
Call sbCbo_Llena_New(cboFormaPago, strSQL, True, True)

strSQL = "select rtrim(cod_caja) as 'IdX', rtrim(Descripcion) as itmx from cajas_definicion  where activa = 1"
Call sbCbo_Llena_New(cboCajas, strSQL, True, True)

cboMov.Clear
cboMov.AddItem "Entradas"
cboMov.AddItem "Salidas"
cboMov.AddItem "TODOS"
cboMov.Text = "TODOS"

vPaso = False

vGrid.MaxRows = 0

dtpRegistroInicio.Value = fxFechaServidor
dtpRegistroCorte.Value = dtpRegistroInicio.Value

Call chkFechas_Click

Call RefrescaTags(Me)
Call Formularios(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next


vGrid.Width = Me.Width - 300
vGrid.Height = Me.Height - (vGrid.top + 600)

imgBanner.Width = Me.Width

End Sub


Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer, curTotal As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbSIFCleanTxtInject(txtCedula)
Call sbSIFCleanTxtInject(txtNombre)
Call sbSIFCleanTxtInject(txtUsuario)
Call sbSIFCleanTxtInject(txtNumDoc)

curTotal = 0

strSQL = "select * " _
       & " From vCaja_AfectacionFormaPago"


If chkFechas.Value = vbUnchecked Then
    strSQL = strSQL & " Where registro_fecha between '" & Format(dtpRegistroInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpRegistroCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Else
    strSQL = strSQL & " Where registro_fecha between '1900/01/01 00:00:00' and dbo.MyGetdate()"
End If

If Len(Trim(txtCedula.Text)) > 0 Then
    strSQL = strSQL & " and Cliente_Identificacion like '%" & txtCedula.Text & "%'"
End If

If Len(Trim(txtNombre.Text)) > 0 Then
    strSQL = strSQL & " and Cliente_Nombre like '%" & txtNombre.Text & "%'"
End If

If Trim(cboFormaPago.Text) <> "TODOS" Then
    strSQL = strSQL & " and COD_FORMA_PAGO in('" & cboFormaPago.ItemData(cboFormaPago.ListIndex) & "','SF')"
End If

If Len(Trim(txtNumDoc.Text)) > 0 Then
    strSQL = strSQL & " and NUM_REFERENCIA like '%" & txtNumDoc.Text & "%'"
End If

If Len(Trim(txtUsuario.Text)) > 0 Then
    strSQL = strSQL & " and Registro_Usuario like '%" & txtUsuario.Text & "%'"
End If

If chkSF_Relacionados.Value = vbUnchecked Then
    strSQL = strSQL & " and cod_forma_pago not in('SF')"
End If

If Trim(cboCajas.Text) <> "TODOS" Then
    strSQL = strSQL & " and COD_CAJA in('" & cboCajas.ItemData(cboCajas.ListIndex) & "')"
End If

If IsNumeric(txtCajaAP.Text) Then
    strSQL = strSQL & " and COD_APERTURA = " & txtCajaAP.Text
End If

Select Case Mid(cboMov.Text, 1, 1)
 Case "E"
     strSQL = strSQL & " and Monto_Aplicado >= 0"

 Case "S"
     strSQL = strSQL & " and Monto_Aplicado < 0"
End Select

Call OpenRecordSet(rs, strSQL)

vGrid.MaxRows = 0


  Do While Not rs.EOF
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
         
   
    For i = 1 To vGrid.MaxCols
      vGrid.col = i
      Select Case i
         Case 1 'Cedula
            vGrid.Text = Trim(rs!Cliente_Identificacion & "")
         Case 2 'Nombre
            vGrid.Text = rs!Cliente_Nombre & ""
         Case 3 'Tipo Doc
            vGrid.Text = rs!TipoDocDesc
         Case 4 'Num Documento
            vGrid.Text = rs!Cod_Transaccion
         Case 5 'Monto
            vGrid.Text = Format(rs!Monto_Doc, "Standard")
         Case 6 'Monto Aplicado
            vGrid.Text = Format(rs!Monto_Aplicado, "Standard")
            
            
         Case 7 'Divisa
            vGrid.Text = rs!cod_Divisa & ""
         Case 8 'Tipo de Cambio
            vGrid.Text = rs!Tipo_Cambio & ""
           
         
         
         Case 9 'Registro Fecha
            vGrid.Text = rs!Registro_Fecha & ""
         Case 10 'Registro Usuario
            vGrid.Text = rs!Registro_Usuario & ""
      
      
         Case 11 'Forma de Pago
            vGrid.Text = rs!FormaPagoDesc & ""
         Case 12 'Referencia
            vGrid.Text = rs!Num_Referencia & ""
         Case 13 'Banco
            vGrid.Text = rs!BancoDesc & ""
         Case 14 'Pagador
            vGrid.Text = rs!EntidadPagoDesc & ""
         Case 15 'Cuenta
            vGrid.Text = rs!cod_cuenta & ""
         Case 16 'Concepto
            vGrid.Text = rs!ConceptoDesc & ""
         Case 17 'Caja
            vGrid.Text = rs!cod_caja & ""
         Case 18 'Apertura
            vGrid.Text = rs!cod_Apertura & ""
     
      
      End Select
    Next i
    
     curTotal = curTotal + rs!Monto_Aplicado
    
     rs.MoveNext
   Loop
rs.Close

txtTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

