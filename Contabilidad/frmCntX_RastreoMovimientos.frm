VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_RastreoMovimientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Rastreo de Movimientos a Cuentas"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16800
   HelpContextID   =   15
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   16800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   12375
      _Version        =   1310723
      _ExtentX        =   21828
      _ExtentY        =   7858
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      _Version        =   1310723
      _ExtentX        =   4683
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboMovimiento 
      Height          =   330
      Left            =   5640
      TabIndex        =   5
      Top             =   600
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboSignos 
      Height          =   330
      Left            =   7560
      TabIndex        =   6
      Top             =   600
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.FlatEdit txtCuentaInicio 
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin XtremeSuiteControls.FlatEdit txtCuentaCorte 
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Top             =   1440
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   315
      Left            =   9000
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   315
      Left            =   9000
      TabIndex        =   17
      Top             =   1440
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin XtremeSuiteControls.FlatEdit txtDesCuentaCorte 
      Height          =   315
      Left            =   3480
      TabIndex        =   21
      Top             =   1440
      Width           =   4095
      _Version        =   1310723
      _ExtentX        =   7223
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDesCuentaInicio 
      Height          =   315
      Left            =   3480
      TabIndex        =   20
      Top             =   1080
      Width           =   4095
      _Version        =   1310723
      _ExtentX        =   7223
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtParametro 
      Height          =   330
      Left            =   9000
      TabIndex        =   22
      Top             =   600
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   5640
      TabIndex        =   23
      Top             =   120
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcionCodigo 
      Height          =   315
      Left            =   6720
      TabIndex        =   24
      Top             =   120
      Width           =   4335
      _Version        =   1310723
      _ExtentX        =   7646
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   12120
      TabIndex        =   25
      Top             =   600
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_RastreoMovimientos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   13440
      TabIndex        =   26
      ToolTipText     =   "Exportar a Excel"
      Top             =   600
      Width           =   615
      _Version        =   1310723
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_RastreoMovimientos.frx":0A1E
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   12120
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   233
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   7680
      TabIndex        =   19
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   7680
      TabIndex        =   18
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tpo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -7680
      TabIndex        =   0
      Top             =   -1080
      Width           =   615
   End
End
Attribute VB_Name = "frmCntX_RastreoMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem


Private Sub sbLimpiaDatos()

txtCodigo = ""
txtDescripcionCodigo = ""
txtCuentaInicio = ""
txtDesCuentaInicio = ""
txtCuentaCorte = ""
txtDesCuentaCorte = ""
txtParametro = ""

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSignos.SetFocus
End Sub

Private Sub cboSignos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtParametro.SetFocus
End Sub

Private Sub cboTipo_Click()
Call sbLimpiaDatos
End Sub

Private Sub btnBuscar_Click()
Dim curDebitos As Currency, curCreditos As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.cod_cuenta, C.cod_cuenta_Mask, C.descripcion,D.tipo_asiento,D.num_asiento" _
       & ",D.fecha_asiento,D.monto_debito,D.monto_credito,D.Documento,D.detalle" _
       & ",D.cod_unidad,D.cod_Centro_Costo,D.cod_Divisa,E.COD_CONTABILIDAD,E.nombre" _
       & " from CntX_Asientos A inner join CntX_Asientos_detalle D" _
       & " on A.COD_CONTABILIDAD = D.COD_CONTABILIDAD and A.tipo_asiento = D.tipo_asiento and A.num_asiento = D.num_asiento" _
       & " inner join CNTX_CONTABILIDADES E on D.COD_CONTABILIDAD = E.COD_CONTABILIDAD" _
       & " inner join CntX_Cuentas C on D.COD_CONTABILIDAD = C.COD_CONTABILIDAD and D.cod_cuenta = C.cod_cuenta" _
       & " where D.fecha_asiento between '" & Format(dtpInicio, "yyyy/mm/dd") & "' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & "' and C.cod_cuenta_Mask between '" & Trim(txtCuentaInicio.Text) _
       & "' and '" & Trim(txtCuentaCorte.Text) & "'"


If IsNumeric(txtParametro.Text) Then
 If CCur(txtParametro.Text) > 0 Then
    Select Case cboMovimiento.Text
      Case "Debitos"
         strSQL = strSQL & " and D.monto_debito " & cboSignos.Text & " " & CCur(txtParametro)
      Case "Creditos"
         strSQL = strSQL & " and D.monto_credito " & cboSignos.Text & " " & CCur(txtParametro)
      Case "Ambos"
         strSQL = strSQL & " and D.monto_debito " & cboSignos.Text & " " & CCur(txtParametro)
         strSQL = strSQL & " or D.monto_credito " & cboSignos.Text & " " & CCur(txtParametro)
    End Select
 End If
End If

If Len(Trim(txtDocumento.Text)) > 0 Then
     strSQL = strSQL & " and D.documento like '%" & txtDocumento.Text & "%'"
End If

If Len(Trim(txtDetalle.Text)) > 0 Then
     strSQL = strSQL & " and D.detalle like '%" & txtDetalle.Text & "%'"
End If

If cboTipo.Text = "Contabilidad Individual" Then
   strSQL = strSQL & " and E.COD_CONTABILIDAD = " & txtCodigo
Else 'CNTX_CONSOLIDA_DEFINICION
   strSQL = strSQL & " and E.COD_CONTABILIDAD in(select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION_DET" _
          & " where cod_consolida = " & txtCodigo & ")"
End If

strSQL = strSQL & " order by D.fecha_asiento,D.cod_cuenta"

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL, 0)
ProgressBarX.Value = 1
ProgressBarX.Max = rs.RecordCount + 1
ProgressBarX.Visible = True
curDebitos = 0
curCreditos = 0

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_cuenta_Mask)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = Format(rs!fecha_asiento, "yyyy-mm-dd")
      itmX.SubItems(3) = UCase(rs!Tipo_Asiento)
      itmX.SubItems(4) = rs!Num_Asiento
      itmX.SubItems(5) = Format(rs!monto_debito, "Standard")
      itmX.SubItems(6) = Format(rs!monto_credito, "Standard")
      itmX.SubItems(7) = rs!Nombre
      itmX.SubItems(8) = rs!Documento
      itmX.SubItems(9) = rs!Detalle
  
  curDebitos = curDebitos + rs!monto_debito
  curCreditos = curCreditos + rs!monto_credito
  
  ProgressBarX.Value = ProgressBarX.Value + 1
  rs.MoveNext
Loop
rs.Close
ProgressBarX.Value = 1
ProgressBarX.Visible = False

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(5) = "--------------------------"
    itmX.SubItems(6) = "--------------------------"


Set itmX = lsw.ListItems.Add(, , "TOTALES")
    itmX.SubItems(5) = Format(curDebitos, "Standard")
    itmX.SubItems(6) = Format(curCreditos, "Standard")

Me.MousePointer = vbDefault
MsgBox "Consulta Finalizada...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaInicio.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub Form_Load()

cboMovimiento.Clear
cboMovimiento.AddItem "Debitos"
cboMovimiento.AddItem "Creditos"
cboMovimiento.AddItem "Ambos"
cboMovimiento.Text = "Debitos"

cboSignos.Clear
cboSignos.AddItem "="
cboSignos.AddItem ">"
cboSignos.AddItem "<"
cboSignos.Text = "="

cboTipo.Clear
cboTipo.AddItem "Contabilidad Individual"
cboTipo.AddItem "Consolidación"
cboTipo.Text = "Contabilidad Individual"


Call sbInicializa

End Sub

Private Sub Form_Resize()
On Error Resume Next

lsw.Width = Me.Width - 250
lsw.Height = Me.Height - (lsw.Top + 400)


End Sub

Private Sub sbInicializa()

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "Cuenta", 2100, vbCenter
lsw.ColumnHeaders.Add , , "Descripción", 3200
lsw.ColumnHeaders.Add , , "Fecha", 1800, vbCenter
lsw.ColumnHeaders.Add , , "Tipo Asiento", 1200, vbCenter
lsw.ColumnHeaders.Add , , "Num. Asiento", 2500, vbCenter
lsw.ColumnHeaders.Add , , "Debitos", 2100, vbRightJustify
lsw.ColumnHeaders.Add , , "Creditos", 2100, vbRightJustify
lsw.ColumnHeaders.Add , , "Empresa", 3200
lsw.ColumnHeaders.Add , , "Documento", 2200
lsw.ColumnHeaders.Add , , "Detalle", 4200


cboMovimiento.Text = "Debitos"
cboSignos.Text = "="
cboTipo.Text = "Contabilidad Individual"

'Cargar la Contabilidad Aqui

txtCodigo = gCntX_Parametros.CodigoConta
txtDescripcionCodigo = gCntX_Parametros.NombreEmpresa

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio

txtParametro = 0


End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcionCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  If cboTipo.Text = "Contabilidad Individual" Then
      gBusquedas.Columna = "COD_CONTABILIDAD"
      gBusquedas.Orden = "COD_CONTABILIDAD"
      gBusquedas.Consulta = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES"
      gBusquedas.Filtro = ""
  Else 'CONSOLIDA_DEFINICION
      gBusquedas.Columna = "cod_consolida"
      gBusquedas.Orden = "cod_consolida"
      gBusquedas.Consulta = "select cod_consolida,descripcion from CNTX_CONSOLIDA_DEFINICION"
      gBusquedas.Filtro = ""
  End If
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcionCodigo.Text = gBusquedas.Resultado2
  txtCodigo.SetFocus
End If
End Sub

Private Sub txtCodigo_LostFocus()
Dim xCodigo As Long

On Error GoTo vError

If cboTipo = "Contabilidad Individual" Then
 xCodigo = txtCodigo

Else
 strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
 Call OpenRecordSet(rs, strSQL, 0)
 xCodigo = rs!COD_CONTABILIDAD
 rs.Close
End If

strSQL = "select * from CNTX_CONTABILIDADES where COD_CONTABILIDAD = " & xCodigo
Call OpenRecordSet(rs, strSQL, 0)
    txtCodigo.Text = xCodigo
    txtDescripcionCodigo.Text = rs!Nombre
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCuentaCorte_KeyDown(KeyCode As Integer, Shift As Integer)
Dim xCodigo As Long

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboMovimiento.SetFocus

If KeyCode = vbKeyF4 Then
  If cboTipo.Text = "Contabilidad Individual" Then
   xCodigo = txtCodigo
  Else
    strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
    Call OpenRecordSet(rs, strSQL, 0)
     xCodigo = rs!COD_CONTABILIDAD
    rs.Close
  End If
  
  gBusquedas.Columna = "cod_cuenta_mask"
  gBusquedas.Orden = "cod_cuenta_mask"
  gBusquedas.Consulta = "select cod_cuenta_mask,descripcion,acepta_movimientos from CntX_Cuentas"
  gBusquedas.Filtro = " and COD_CONTABILIDAD = " & xCodigo
  frmBusquedas.Show vbModal
  txtCuentaCorte = gBusquedas.Resultado
  txtDesCuentaCorte.Text = gBusquedas.Resultado2

  txtCuentaCorte.SetFocus
End If



End Sub

Private Sub txtCuentaCorte_LostFocus()

Dim xCodigo As Long

On Error GoTo vError

  If cboTipo.Text = "Contabilidad Individual" Then
   xCodigo = txtCodigo
  Else
    strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
    Call OpenRecordSet(rs, strSQL, 0)
     xCodigo = rs!COD_CONTABILIDAD
    rs.Close
  End If
  
  strSQL = "select descripcion from CntX_cuentas where COD_CONTABILIDAD = " & xCodigo _
         & " and cod_cuenta = '" & Replace(txtCuentaCorte.Text, "-", "") & "'"
  Call OpenRecordSet(rs, strSQL, 0)
  txtDesCuentaCorte = rs!Descripcion & ""
  rs.Close
  
vError:

End Sub

Private Sub txtCuentaInicio_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xCodigo As Long

On Error GoTo vError

  If cboTipo.Text = "Contabilidad Individual" Then
   xCodigo = txtCodigo
  Else
    strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
    Call OpenRecordSet(rs, strSQL, 0)
     xCodigo = rs!COD_CONTABILIDAD
    rs.Close
  End If
  
  strSQL = "select descripcion from CntX_cuentas where COD_CONTABILIDAD = " & xCodigo _
         & " and cod_cuenta = '" & Replace(txtCuentaInicio.Text, "-", "") & "'"
  Call OpenRecordSet(rs, strSQL, 0)
  txtDesCuentaInicio = rs!Descripcion & ""
  rs.Close
  
vError:

End Sub

Private Sub txtCuentaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
Dim xCodigo As Long

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesCuentaCorte.SetFocus

If KeyCode = vbKeyF4 Then
  If cboTipo.Text = "Contabilidad Individual" Then
   xCodigo = txtCodigo
  Else
    strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
    Call OpenRecordSet(rs, strSQL, 0)
     xCodigo = rs!COD_CONTABILIDAD
    rs.Close
  End If
  
    gBusquedas.Columna = "cod_Cuenta_Mask"
    gBusquedas.Orden = "cod_Cuenta_Mask"
    gBusquedas.Consulta = "select cod_Cuenta_Mask,descripcion,acepta_movimientos from CntX_cuentas"
    gBusquedas.Filtro = " and COD_CONTABILIDAD = " & xCodigo
    frmBusquedas.Show vbModal
    txtCuentaInicio = gBusquedas.Resultado
    txtDesCuentaInicio.Text = gBusquedas.Resultado2
  
    txtCuentaInicio.SetFocus
End If

End Sub

Private Sub txtDescripcionCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus

If KeyCode = vbKeyF4 Then
  If cboTipo.Text = "Contabilidad Individual" Then
      gBusquedas.Columna = "nombre"
      gBusquedas.Orden = "nombre"
      gBusquedas.Consulta = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES"
      gBusquedas.Filtro = ""
  Else 'CNTX_CONSOLIDA_DEFINICION
      gBusquedas.Columna = "descripcion"
      gBusquedas.Orden = "descripcion"
      gBusquedas.Consulta = "select cod_consolida,descripcion from CNTX_CONSOLIDA_DEFINICION"
      gBusquedas.Filtro = ""
  End If
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtCodigo.SetFocus

End If
End Sub


Private Sub txtDesCuentaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim xCodigo As Long

If KeyCode = vbKeyF4 Then
  If cboTipo.Text = "Contabilidad Individual" Then
   xCodigo = txtCodigo
  Else
    strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
    Call OpenRecordSet(rs, strSQL, 0)
     xCodigo = rs!COD_CONTABILIDAD
    rs.Close
  End If
  
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_Cuenta_Mask,descripcion,acepta_movimientos from CntX_cuentas"
  gBusquedas.Filtro = " and COD_CONTABILIDAD = " & xCodigo
  frmBusquedas.Show vbModal
  txtCuentaInicio = gBusquedas.Resultado
  txtDesCuentaInicio.Text = gBusquedas.Resultado2
  txtCuentaInicio.SetFocus
End If

End Sub


Private Sub txtDesCuentaCorte_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim xCodigo As Long

If KeyCode = vbKeyF4 Then
  If cboTipo.Text = "Contabilidad Individual" Then
   xCodigo = txtCodigo
  Else
    strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & txtCodigo
    Call OpenRecordSet(rs, strSQL, 0)
     xCodigo = rs!COD_CONTABILIDAD
    rs.Close
  End If
  
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_Cuenta_Mask,descripcion,acepta_movimientos from CntX_cuentas"
  gBusquedas.Filtro = " and COD_CONTABILIDAD = " & xCodigo
  frmBusquedas.Show vbModal
  txtCuentaCorte = gBusquedas.Resultado
  txtDesCuentaCorte.Text = gBusquedas.Resultado2
  txtCuentaCorte.SetFocus
End If

End Sub


Private Sub txtParametro_GotFocus()
On Error GoTo vError
txtParametro = CCur(txtParametro)
vError:
End Sub

Private Sub txtParametro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnBuscar.SetFocus
End Sub

Private Sub txtParametro_LostFocus()
On Error GoTo vError
    txtParametro = Format(CCur(txtParametro), "Standard")
vError:
End Sub
