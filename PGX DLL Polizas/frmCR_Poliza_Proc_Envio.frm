VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_Poliza_Proc_Envio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Pólizas: Generación de Archivos"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7695
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   14175
      _Version        =   1572864
      _ExtentX        =   25003
      _ExtentY        =   13573
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
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   330
      Left            =   8040
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   10080
      TabIndex        =   5
      Top             =   1080
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Envio.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnGenerar 
      Height          =   375
      Left            =   10560
      TabIndex        =   6
      ToolTipText     =   "Generar"
      Top             =   1080
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Envio.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnPrevista 
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      ToolTipText     =   "Prevista"
      Top             =   1080
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Prevista"
      BackColor       =   16777215
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
      Picture         =   "frmCR_Poliza_Proc_Envio.frx":0E19
   End
   Begin XtremeSuiteControls.FlatEdit txtRegistros 
      Height          =   315
      Left            =   13080
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPoliza 
      Height          =   330
      Left            =   840
      TabIndex        =   11
      Top             =   1080
      Width           =   4215
      _Version        =   1572864
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   8040
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   14175
      _Version        =   1572864
      _ExtentX        =   25003
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registros:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   12240
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Corte:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Póliza.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Generación de archivos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Polizas de Vivienda y Prendario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmCR_Poliza_Proc_Envio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, vTipo As String

Private Sub btnBuscar_Click()

'spPoliza_Incendio(@Poliza varchar(10), @Corte datetime, @Beneficiarios smallint = 1, @Usuario varchar(30)= ''
'        , @Movimiento varchar(5) = 'T', @Cedula varchar(20) = Null)

strSQL = "exec spPoliza_Incendio '" & cboPoliza.ItemData(cboPoliza.ListIndex) & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
       & "', 0, '" & glogon.Usuario & "', 'T'"
Call OpenRecordSet(rs, strSQL)

With lsw.ColumnHeaders
    .Clear
  Select Case vTipo
    Case "V"  'Vida
            .Add , , "Corte", 2000
            .Add , , "PRIMER_NOMB", 1500
            .Add , , "APELLIDO_PAT", 1500
            .Add , , "APELLIDO_MAT", 1500
            .Add , , "SEXO", 1000, vbCenter
            .Add , , "FECHA_NACIMIENTO", 1500, vbCenter
            .Add , , "NUMERO_CEDULA", 1500, vbCenter
            .Add , , "MONTO_ASEGURADO", 1800, vbRightJustify
            .Add , , "NUMERO_DE_OPERACION", 1500, vbCenter
            .Add , , "TIPO_POLIZA", 1500, vbCenter
    
    Case "I" 'Incendio
        .Add , , "Corte", 2000
        .Add , , "Identificación", 2000
        .Add , , "Primer Apellido", 2000
        .Add , , "Segundo Apellido", 2000
        .Add , , "Primer Nombre", 2000
        .Add , , "Segundo Nombre", 2000
        .Add , , "Teléfono", 2000
        .Add , , "Correo Electrónico", 2000
        .Add , , "No. Folio", 2000
        .Add , , "Provincia", 2000
        .Add , , "Cantón", 2000
        .Add , , "Distrito", 2000
        .Add , , "Dirección Completa", 3000
        .Add , , "Monto del Crédito", 1800, vbRightJustify
        .Add , , "Monto de Construcción", 1800, vbRightJustify
        .Add , , "No. Operación", 2000
        .Add , , "No. Finca", 2000
        .Add , , "Tipo", 2000
        
        Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!Corte)
             itmX.SubItems(1) = rs!Cedula
             itmX.SubItems(2) = rs!Apellido_1
             itmX.SubItems(3) = rs!Apellido_2
             itmX.SubItems(4) = rs!Nombre_1
             itmX.SubItems(5) = rs!Nombre_2
             itmX.SubItems(6) = rs!Telefono
             itmX.SubItems(7) = rs!Email
             itmX.SubItems(8) = rs!Folio & ""
             itmX.SubItems(9) = rs!Provincia
             itmX.SubItems(10) = rs!Canton
             itmX.SubItems(11) = rs!Distrito & ""
             itmX.SubItems(12) = rs!Direccion & ""
             itmX.SubItems(13) = Format(rs!Monto, "Standard")
             itmX.SubItems(14) = Format(rs!ValorConstruccion, "Standard")
             itmX.SubItems(15) = rs!Id_Solicitud
             itmX.SubItems(16) = rs!NumeroFinca
             itmX.SubItems(17) = rs!Movimiento
         
         rs.MoveNext
        Loop
        
    Case "PC" 'MAC - INS
  End Select
  
End With

rs.Close

txtRegistros.Text = Format(lsw.ListItems.Count, "###,##0")

End Sub

Private Sub btnPrevista_Click()
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

Private Sub cboPoliza_Click()
Dim strSQL As String

If vPaso Then Exit Sub
If cboPoliza.ListCount < 0 Then Exit Sub


'strSQL = "select Prov.COD_PROVEEDOR as 'IdX', Prov.DESCRIPCION  as 'ItmX'" _
'       & " from CRD_CATALOGO_POLIZAS Cp" _
'       & "   inner join CRD_POLIZAS_ASEGURADORAS Pa  on Cp.COD_ASEGURADORA = Pa.COD_ASEGURADORA" _
'       & "   inner join CXP_PROVEEDORES Prov on Pa.COD_PROVEEDOR = Prov.COD_PROVEEDOR" _
'       & " Where Cp.COD_POLIZA = '" & cboPoliza.ItemData(cboPoliza.ListIndex) & "'"
'vPaso = True
'Call sbCbo_Llena_New(cboProveedor, strSQL, False, True)
'vPaso = False

strSQL = "select case when Pg.TIPO_APLICACION  in('PINC') then 'I'" _
       & "      when Pg.TIPO_APLICACION  in('PPC') then 'PC'" _
       & "     Else  'V' end as 'TIPO'" _
       & ", Cp.COD_POLIZA, Cp.CODIGO_RETENCION" _
       & "  from CRD_CATALOGO_POLIZAS Cp inner join POLIZAS_GRUPO Pg on Cp.ID_POLIZA_GRUPO = Pg.ID_POLIZA_GRUPO" _
       & " Where Cp.COD_POLIZA = '" & cboPoliza.ItemData(cboPoliza.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

With lsw.ColumnHeaders
    .Clear
  vTipo = rs!Tipo
  Select Case rs!Tipo
    Case "V"  'Vida
            .Add , , "PRIMER_NOMB", 1500
            .Add , , "APELLIDO_PAT", 1500
            .Add , , "APELLIDO_MAT", 1500
            .Add , , "SEXO", 1000, vbCenter
            .Add , , "FECHA_NACIMIENTO", 1500, vbCenter
            .Add , , "NUMERO_CEDULA", 1500, vbCenter
            .Add , , "MONTO_ASEGURADO", 1800, vbRightJustify
            .Add , , "NUMERO_DE_OPERACION", 1500, vbCenter
            .Add , , "TIPO_POLIZA", 1500, vbCenter
    
    Case "I" 'Incendio
        .Add , , "Identificación", 2000
        .Add , , "Primer Apellido", 2000
        .Add , , "Segundo Apellido", 2000
        .Add , , "Primer Nombre", 2000
        .Add , , "Segundo Nombre", 2000
        .Add , , "Teléfono", 2000
        .Add , , "Correo Electrónico", 2000
        .Add , , "Teléfono", 2000
        .Add , , "No. Folio", 2000
        .Add , , "Provincia", 2000
        .Add , , "Cantón", 2000
        .Add , , "Distrito", 2000
        .Add , , "Dirección Completa", 3000
        .Add , , "Monto del Crédito", 1800, vbRightJustify
        .Add , , "Monto de Construcción", 1800, vbRightJustify
        .Add , , "No. Operación", 2000
        .Add , , "No. Finca", 2000
        .Add , , "Tipo", 2000
    Case "PC" 'MAC - INS
  End Select
End With

rs.Close

End Sub

Private Sub Form_Load()
vModulo = 11

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Dim vFecha As Date, vFechaInicial As Date, vProceso As Currency, i As Integer

vFecha = fxFechaServidor
vFechaInicial = vFecha


vPaso = True
    strSQL = "select COD_POLIZA as 'IdX', DESCRIPCION as 'ItmX' From CRD_CATALOGO_POLIZAS"
    Call sbCbo_Llena_New(cboPoliza, strSQL, False, True)
vPaso = False

dtpCorte.Value = vFecha

vFecha = DateAdd("m", -10, vFecha)
vProceso = Format(vFecha, "yyyymm")
For i = 1 To 12
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboProceso.AddItem CStr(vProceso)
Next i

vProceso = Format(vFechaInicial, "yyyymm")
cboProceso.Text = vProceso

Call cboPoliza_Click


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
lsw.Width = Me.Width - 200
lsw.Height = Me.Height - (lsw.Top + 450)
ProgressBarX.Width = lsw.Width

End Sub

Private Sub txtPoliza_Codigo_Change()

End Sub
