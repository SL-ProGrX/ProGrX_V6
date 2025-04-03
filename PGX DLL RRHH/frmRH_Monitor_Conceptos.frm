VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmRH_Monitor_Conceptos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "RRHH: Monitor de Conceptos"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17160
   LinkTopic       =   "Form5"
   ScaleHeight     =   10425
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3615
      Left            =   0
      TabIndex        =   19
      Top             =   2160
      Width           =   15015
      _Version        =   1441792
      _ExtentX        =   26485
      _ExtentY        =   6376
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Consulta"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Resumen"
      Item(1).ControlCount=   0
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3375
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   12975
         _Version        =   524288
         _ExtentX        =   22886
         _ExtentY        =   5953
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
         MaxCols         =   11
         SpreadDesigner  =   "frmRH_Monitor_Conceptos.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin XtremeSuiteControls.CheckBox chkNominaDetalla 
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Detallar por Nómina"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   868
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
      Picture         =   "frmRH_Monitor_Conceptos.frx":07F0
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   495
      Left            =   10560
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2143
      _ExtentY        =   873
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
      Picture         =   "frmRH_Monitor_Conceptos.frx":0EF0
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   312
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   5412
      _Version        =   1441792
      _ExtentX        =   9551
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1452
      _Version        =   1441792
      _ExtentX        =   2561
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
      Left            =   4320
      TabIndex        =   4
      Top             =   480
      Width           =   1452
      _Version        =   1441792
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
      Height          =   312
      Left            =   2520
      TabIndex        =   11
      Top             =   1080
      Width           =   4332
      _Version        =   1441792
      _ExtentX        =   7641
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConcepto 
      Height          =   312
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Width           =   1092
      _Version        =   1441792
      _ExtentX        =   1926
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEntidadDesc 
      Height          =   312
      Left            =   2520
      TabIndex        =   13
      Top             =   1440
      Width           =   4332
      _Version        =   1441792
      _ExtentX        =   7641
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEntidad 
      Height          =   312
      Left            =   1440
      TabIndex        =   14
      Top             =   1440
      Width           =   1092
      _Version        =   1441792
      _ExtentX        =   1926
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   2520
      TabIndex        =   15
      Top             =   1800
      Width           =   4335
      _Version        =   1441792
      _ExtentX        =   7646
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
      _Version        =   1441792
      _ExtentX        =   1931
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo_ProGrX 
      Height          =   495
      Left            =   11760
      TabIndex        =   21
      Top             =   1440
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Archivo Plano"
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
      Picture         =   "frmRH_Monitor_Conceptos.frx":17C1
      ImageAlignment  =   4
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Entidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   732
   End
   Begin VB.Image imgBanner 
      Height          =   2190
      Left            =   0
      Picture         =   "frmRH_Monitor_Conceptos.frx":1EDA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13125
   End
End
Attribute VB_Name = "frmRH_Monitor_Conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnArchivo_ProGrX_Click()
Dim pCedula As String, pMonto As String, pCodigo As String, pNombre As String
Dim fn, vCadena As String
Dim vArchivo As String, vTempo As String, vRuta As String, vFile As String

Dim i As Long


On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\RRHH\"
vRuta = SIFGlobal.DirectorioDeResultados & "\RRHH\"

vArchivo = "RRHH-" & txtEntidad.Text & "_" & Format(dtpInicio.Value, "yyyymmdd") _
        & "_" & Format(dtpCorte.Value, "yyyymmdd") & ".txt"

vTempo = vRuta & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


On Error GoTo vError

fn = FreeFile

Open vTempo For Output As #fn  ' Create file name.

Me.MousePointer = vbHourglass

With vGrid

For i = 1 To .MaxRows
    .Row = i
    .col = 2
    pCedula = Trim(.Text)
    .col = 3
    pNombre = Trim(.Text)
    .col = 4
    pMonto = Trim(.Text)
    .col = 5
    pCodigo = Trim(.Text)
    
    pMonto = Replace(pMonto, ",", "")
    pMonto = Replace(pMonto, ".", "")
    
    
    pCedula = SIFGlobal.fxStringRelleno(pCedula, "D", " ", 20)
    pCodigo = SIFGlobal.fxStringRelleno(pCodigo, "D", " ", 10)
    pMonto = SIFGlobal.fxStringRelleno(pMonto, "I", " ", 10)
    pNombre = SIFGlobal.fxStringRelleno(pNombre, "D", " ", 30)
 
    vCadena = pCedula + " " + pCodigo + "M " + pMonto + " " + pNombre
 
 If Len(RTrim(vCadena)) > 0 Then
    Print #fn, vCadena
 End If

Next i


End With

Close #fn

  
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

If chkNominaDetalla.Value = xtpChecked Then
    strSQL = " select C.EMPLEADO_ID, C.IDENTIFICACION, C.NOMBRE_COMPLETO, SUM(C.MONTO) AS 'MONTO', C.COD_CONCEPTO, C.CONCEPTO_DESC" _
           & " ,C.COD_NOMINA, C.NOMINA_NUM, CONVERT(VARCHAR(10), C.FECHA_INICIO, 23) as 'FECHA_INICIO', CONVERT(VARCHAR(10), C.FECHA_CORTE, 23)  as 'FECHA_CORTE'" _
           & " ,C.TIPO_CONCEPTO_DESC " _
           & " from vRH_Nomina_Detalle_Conceptos C" _
           & " where C.FECHA_CORTE between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
           & " and C.cod_Nomina = '" & cboNomina.ItemData(cboNomina.ListIndex) & "'"
Else
    strSQL = " select C.EMPLEADO_ID, C.IDENTIFICACION, C.NOMBRE_COMPLETO, SUM(C.MONTO) AS 'MONTO', C.COD_CONCEPTO, C.CONCEPTO_DESC" _
           & " ,C.COD_NOMINA, convert(varchar(30), min(C.NOMINA_NUM)) + '..' + convert(varchar(30), max(C.NOMINA_NUM)) as 'NOMINA_NUM'" _
           & " , CONVERT(VARCHAR(10), min(C.FECHA_INICIO), 23) AS 'FECHA_INICIO', CONVERT(VARCHAR(10), max(C.FECHA_CORTE), 23)   as 'FECHA_CORTE'" _
           & " ,C.TIPO_CONCEPTO_DESC " _
           & " from vRH_Nomina_Detalle_Conceptos C" _
           & " where C.FECHA_CORTE between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
           & " and C.cod_Nomina = '" & cboNomina.ItemData(cboNomina.ListIndex) & "'"
End If

If txtConcepto.Text <> "" Then
    strSQL = strSQL & " And C.cod_concepto= '" & txtConcepto.Text & "'"
End If

If txtEmpleadoId.Text <> "" Then
    strSQL = strSQL & " And C.Empleado_Id = '" & txtEmpleadoId.Text & "'"
End If


If txtEntidad.Text <> "" Then
    strSQL = strSQL & " And C.cod_concepto in(SELECT COD_CONCEPTO  FROM RH_CONCEPTOS WHERE COD_ER = '" & txtEntidad.Text & "')"
End If



If chkNominaDetalla.Value = xtpChecked Then
    strSQL = strSQL & " group by C.EMPLEADO_ID, C.IDENTIFICACION, C.NOMBRE_COMPLETO, C.COD_CONCEPTO, C.CONCEPTO_DESC" _
           & ",C.COD_NOMINA, C.NOMINA_NUM, C.FECHA_INICIO, C.FECHA_CORTE, C.TIPO_CONCEPTO_DESC"

Else
    strSQL = strSQL & " group by C.EMPLEADO_ID, C.IDENTIFICACION, C.NOMBRE_COMPLETO, C.COD_CONCEPTO, C.CONCEPTO_DESC" _
           & ",C.COD_NOMINA, C.TIPO_CONCEPTO_DESC"

End If

Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExportar_Click()
 Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols

    
    vHeaders.Headers(1) = "Empleado Id"
    vHeaders.Headers(2) = "Identificación"
    vHeaders.Headers(3) = "Nombre"
    vHeaders.Headers(4) = "Monto"
    vHeaders.Headers(5) = "Concepto"
    vHeaders.Headers(6) = "Descripción"
    vHeaders.Headers(7) = "Nómina"
    vHeaders.Headers(8) = "Nómina Id"
    vHeaders.Headers(9) = "Fecha Inicio"
    vHeaders.Headers(10) = "Fecha Corte"
    vHeaders.Headers(11) = "Tipo"

If txtEntidad.Text <> "" Then
 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Nomina_Concepto_Entidad_" & txtEntidad.Text)
Else
 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Nomina_Concepto_" & txtConcepto.Text)
End If
End Sub

Private Sub Form_Load()
vPaso = True
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)
vPaso = False

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width


tcMain.Width = Me.Width - 150
tcMain.Height = Me.Height - (tcMain.Top + 500)


vGrid.Width = tcMain.Width - 50
vGrid.Height = tcMain.Height - (vGrid.Top + 150)


End Sub

Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_CONCEPTO"
    gBusquedas.Orden = "COD_CONCEPTO"
    gBusquedas.Consulta = "select COD_CONCEPTO, DESCRIPCION FROM RH_CONCEPTOS"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    txtConcepto.Text = gBusquedas.Resultado
    txtConceptoDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub sbBusca()
   gBusquedas.Convertir = "N"
   gBusquedas.Col1Name = "Empleado Id"
   gBusquedas.Col2Name = "Persona Id"
   gBusquedas.Col3Name = "Nombre"
   gBusquedas.Columna = "Empleado_ID"
   gBusquedas.Orden = "Empleado_ID"
   gBusquedas.Consulta = "Select Empleado_ID,Identificacion,Nombre_Completo From Rh_Personas"
   
   gBusquedas.Filtro = " and ESTADO_PERSONA = 'A'"
   
   frmBusquedas.Show vbModal
   
   txtEmpleadoId.Text = gBusquedas.Resultado
'   txtIdentificacion.Text = Trim(gBusquedas.Resultado2)
   txtNombre.Text = gBusquedas.Resultado3
       
End Sub


Private Sub txtEmpleadoId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_ER"
    gBusquedas.Orden = "COD_ER"
    gBusquedas.Consulta = "select COD_ER, NOMBRE FROM RH_ENTIDADES_RELACIONADAS"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    txtEntidad.Text = gBusquedas.Resultado
    txtEntidadDesc.Text = gBusquedas.Resultado2
End If

End Sub
