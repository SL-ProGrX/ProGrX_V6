VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCxC_BitacoraEspecial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácora Especial: Cuentas por Cobrar"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox fraRevision 
      Height          =   855
      Left            =   8880
      TabIndex        =   5
      Top             =   -120
      Width           =   4815
      _Version        =   1572864
      _ExtentX        =   8488
      _ExtentY        =   1503
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cboRevision 
         Height          =   312
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.CheckBox chkRevision 
         Height          =   252
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   4332
      End
   End
   Begin XtremeSuiteControls.GroupBox gbToolBar 
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   -120
      Width           =   4695
      _Version        =   1572864
      _ExtentX        =   8281
      _ExtentY        =   1508
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   492
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCXC_BitacoraEspecial.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   492
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCXC_BitacoraEspecial.frx":0A1E
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   492
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Exportar"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCXC_BitacoraEspecial.frx":11DA
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6255
      Left            =   4200
      TabIndex        =   0
      Top             =   915
      Width           =   12735
      _Version        =   524288
      _ExtentX        =   22463
      _ExtentY        =   11033
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
      MaxCols         =   496
      SpreadDesigner  =   "frmCXC_BitacoraEspecial.frx":19B6
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   7212
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3972
      _Version        =   1572864
      _ExtentX        =   7006
      _ExtentY        =   12721
      _StockProps     =   79
      Caption         =   "Filtros"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswMovimientos 
         Height          =   4452
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
         _ExtentY        =   7853
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
         Checkboxes      =   -1  'True
         View            =   3
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Height          =   312
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   252
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkUsuarios 
         Height          =   252
         Left            =   2760
         TabIndex        =   19
         Top             =   1080
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkMovimientos 
         Height          =   252
         Left            =   2760
         TabIndex        =   20
         Top             =   1920
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   1320
         TabIndex        =   21
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1080
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   1320
         TabIndex        =   22
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1440
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1572
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Movimientos ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1092
      End
   End
End
Attribute VB_Name = "frmCxC_BitacoraEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "Revisado?"
    vHeaders.Headers(2) = "Usuario"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Operación"
    vHeaders.Headers(6) = "Concepto"
    vHeaders.Headers(7) = "Detalle"
    vHeaders.Headers(8) = "Identificacion"
    vHeaders.Headers(9) = "Nombre"
    vHeaders.Headers(10) = "Notas"
    vHeaders.Headers(11) = "Revisado por"
    vHeaders.Headers(12) = "Revisado Fecha"
Call sbSIFGridExportar(vGrid, vHeaders, "CxC_BitacoraEspecial")
    
'Select Case ButtonMenu.Key
'  Case "Excel"
'Call sbSIFGridExportar(vGrid, vHeaders, "CxC_BitacoraEspecial")
'  Case "HTML"
'      Call sbSIFGridExportar(vGrid, vHeaders, "CxC_BitacoraEspecial", "HTML")
'End Select
End Sub

Private Sub btnInforme_Click()
        vGrid.PrintHeader = "Cuentas por Cobrar: Bitácora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Sub

Private Sub cboRevision_Click()
If cboRevision.ListCount = 0 Then Exit Sub
Call sbBuscar
End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub chkMovimientos_Click()
Dim i As Integer

For i = 1 To lswMovimientos.ListItems.Count
  lswMovimientos.ListItems.Item(i).Checked = chkMovimientos.Value
Next i

End Sub


Private Sub chkRevision_Click()
If chkRevision.Value = vbChecked Then
   txtUsuario.BackColor = cboRevision.BackColor
Else
   txtUsuario.BackColor = vbWhite
End If
End Sub

Private Sub chkUsuarios_Click()
 If chkUsuarios.Value = vbChecked Then
   txtUsuario.Enabled = False
 Else
   txtUsuario.Enabled = True
   txtUsuario = "(Presione F4)"
 End If
  
End Sub

Private Sub sbBuscar()
Dim rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass
   
vGrid.MaxRows = 0
vGrid.MaxCols = 12

vPaso = True

Call OpenRecordSet(rs, fxSQL)

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!BITACORA_ID
  
  vGrid.Col = 2
  vGrid.Text = rs!Usuario
  vGrid.Col = 3
  vGrid.Text = rs!fecha 'Format(rs!Fecha, "dd/mm/yyyy")
  vGrid.Col = 4
  vGrid.Text = rs!MovimientoDesc
  vGrid.Col = 5
  vGrid.Text = CStr(rs!Operacion)
  vGrid.Col = 6
  vGrid.Text = rs!cod_Concepto
  vGrid.Col = 7
  vGrid.Text = rs!Detalle
  vGrid.Col = 8
  vGrid.Text = Trim(rs!Cedula)
  vGrid.Col = 9
  vGrid.Text = rs!Nombre
  
  vGrid.Col = 10
  vGrid.Text = rs!Notas & ""
  
  vGrid.Col = 11
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.Col = 12
  vGrid.Text = rs!Revisado_Fecha & ""
  
  
  vGrid.RowHeight(vGrid.Row) = vGrid.MaxTextRowHeight(vGrid.Row)
  rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


vModulo = 31

vGrid.AppearanceStyle = fxGridStyle

lswMovimientos.ColumnHeaders.Add , , "", 3200
lswMovimientos.HideColumnHeaders = True
lswMovimientos.Checkboxes = True




cboRevision.AddItem "TODOS"
cboRevision.AddItem "Pendientes"
cboRevision.AddItem "Revisados"
cboRevision.Text = "TODOS"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

lswMovimientos.ListItems.Clear
strSQL = "select MOVIMIENTO,DESCRIPCION from US_MOVIMIENTOS_BE WHERE MODULO = " & vModulo & " ORDER BY MOVIMIENTO"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswMovimientos.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!Movimiento
     itmX.Checked = chkMovimientos.Value
 rs.MoveNext
Loop
rs.Close


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

gbFiltros.Height = Me.Height

vGrid.Width = Me.Width - (520 + vGrid.Left)
vGrid.Height = Me.Height - (vGrid.top + 600)

lswMovimientos.Height = Me.Height - (lswMovimientos.top + 750)

End Sub


Private Function fxSQL() As String
Dim strSQL As String
Dim vCadena As String, i As Integer


strSQL = "select C.*,R.cod_concepto, S.cedula,S.nombre,M.Descripcion as MovimientoDesc,case when C.revisado_fecha is null then 0 else 1 end as 'Revisado'" _
       & " from CXC_BITACORA_ESPECIAL C inner join CXC_CUENTAS R on C.OPERACION = R.OPERACION" _
       & " inner join CXC_PERSONAS S on S.cedula = R.cedula" _
       & " inner join US_MOVIMIENTOS_BE M on C.Movimiento = M.Movimiento" _
       & " Where M.Modulo = " & vModulo

If Len(Trim(txtCedula.Text)) > 0 Then
  strSQL = strSQL & " and S.cedula like '%" & txtCedula.Text & "%'"
End If
       
If chkFechas.Value = vbUnchecked Then
   If chkRevision.Value = vbChecked Then
        strSQL = strSQL & " and C.Revisado_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
   Else
        strSQL = strSQL & " and C.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
   End If
End If

'Lista de Tipos de Movimientos
vCadena = " and C.movimiento in('"
For i = 1 To lswMovimientos.ListItems.Count
  If lswMovimientos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswMovimientos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"

If chkUsuarios.Value = vbUnchecked Then
  If txtUsuario <> "" And txtUsuario <> "(Presione F4)" Then
     If chkRevision.Value = vbChecked Then
             strSQL = strSQL & " and C.Revisado_Usuario = '" & txtUsuario & "'"
     Else
             strSQL = strSQL & " and C.Usuario = '" & txtUsuario & "'"
     End If
  End If
End If

Select Case Mid(cboRevision.Text, 1, 1)
   Case "P" 'Pendientes
        strSQL = strSQL & " and C.Revisado_Fecha is null"
   Case "R" 'Revisados
        strSQL = strSQL & " and C.Revisado_Fecha is not null"
   Case "T" 'Todos
End Select

If chkRevision.Value = vbChecked Then
    strSQL = strSQL & " order by C.Revisado_fecha"
Else
    strSQL = strSQL & " order by C.fecha"
End If

fxSQL = strSQL

 
End Function





Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select Cedula,Nombre From CXC_PERSONAS"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If
End Sub

Private Sub GroupBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Or Col > 1 Or Not fraRevision.Enabled Then Exit Sub
 
vGrid.Row = Row
vGrid.Col = 1
If vGrid.Value = vbChecked Then
   strSQL = "update CXC_BITACORA_ESPECIAL set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = dbo.MyGetdate()" _
          & " where BITACORA_ID = " & vGrid.CellTag
   Call ConectionExecute(strSQL)

   vGrid.Col = 11
   vGrid.Text = glogon.Usuario
   vGrid.Col = 12
   vGrid.Text = Date
   
End If
End Sub


