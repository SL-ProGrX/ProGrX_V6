VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmRH_Vacaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Registro de Disfrute de Vacaciones"
   ClientHeight    =   5085
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9300
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8640
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   6376
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton rbAccion 
         Height          =   252
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         Top             =   1680
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Disfrutar"
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
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   612
         Left            =   7440
         TabIndex        =   1
         Top             =   3000
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmRH_Vacaciones.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   6852
         _Version        =   1572864
         _ExtentX        =   12091
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   2160
         TabIndex        =   3
         Top             =   600
         Width           =   6852
         _Version        =   1572864
         _ExtentX        =   12086
         _ExtentY        =   1609
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaI 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   2160
         TabIndex        =   5
         Top             =   3120
         Width           =   2655
         _Version        =   1572864
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaC 
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtDias 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         ToolTipText     =   "Dias a Disfrutar"
         Top             =   2640
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiasDisponibles 
         Height          =   315
         Left            =   3480
         TabIndex        =   19
         ToolTipText     =   "Días Disponibles"
         Top             =   2640
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbAccion 
         Height          =   252
         Index           =   1
         Left            =   3840
         TabIndex        =   21
         Top             =   1680
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Liquidar"
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
         Appearance      =   16
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Días"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   2640
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   4
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   1080
         TabIndex        =   6
         Top             =   3120
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   2160
      TabIndex        =   10
      Top             =   480
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   312
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   3960
      TabIndex        =   15
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Id. Empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   2160
      TabIndex        =   13
      Top             =   240
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmRH_Vacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim mNomina As String



Private Sub btnAplicar_Click()
If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

Dim Boleta As String, LiquidaId As Integer

'Validacion
If dtpFechaI.Value > dtpFechaC.Value Then
    MsgBox "Error en Rango de Fechas!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtDias.Text) Then
    MsgBox "Dias de Vacaciones Inválido!", vbExclamation
    Exit Sub
End If

If CCur(txtDias.Text) < 0 Then
    MsgBox "Dias de Vacaciones Inválido!", vbExclamation
    Exit Sub
End If

If Not fxRH_Vacaciones_Valida(mNomina, txtEmpleadoId.Text, dtpFechaI.Value, dtpFechaC.Value) Then
    MsgBox "Existe conflicto de fechas de disfrute con alguna otra boleta procesada o con una Nómina ya ejecutada!", vbExclamation
    Exit Sub
End If


On Error GoTo vError

'TODO: Validar Sobre Giro en Vacaciones

'spRH_Vacaciones_Registro(@EmpleadoId varchar(20), @Tipo varchar(10), @Notas varchar(1000), @Usuario varchar(30)
'                , @Salida datetime, @Entrada datetime, @D_Disfrutados int, @D_Disponibles dec(10,4), @LiquidaID smallint
'                , @Estado char(1) = 'S', @AutorizaId varchar(30) = Null
'                , @AppCod varchar(30) = 'ProGrX' )


Dim pAutorizador As String

If Mid(cboEstado.Text, 1, 1) = "S" Then
  pAutorizador = "Null"
Else
  pAutorizador = "Null"
End If

strSQL = "exec spRH_Vacaciones_Registro '" & txtEmpleadoId.Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) _
        & "','" & txtNotas.Text & "','" & glogon.Usuario & "'" _
        & ",'" & Format(dtpFechaI.Value, "yyyy/mm/dd") & " 00:00:00'" _
        & ",'" & Format(dtpFechaC.Value, "yyyy/mm/dd") & " 23:59:59'" _
        & "," & CInt(txtDias.Text) & "," & CCur(txtDiasDisponibles.Text) & "," & LiquidaId _
        & ",'" & Mid(cboEstado.Text, 1, 1) & "'," & pAutorizador & ",'ProGrX'"
Call OpenRecordSet(rs, strSQL)
    Boleta = rs!BoletaId
rs.Close

'Print Boleta
Call sbBoleta_Vacaciones(Boleta)

MsgBox "Vacaciones registradas satisfactoriamente!", vbInformation


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub cboTipo_Click()
If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select REQUIERE_AUTORIZACION,PERMITE_LIQUIDACION" _
       & " from RH_VACACIONES_TIPOS " _
       & " WHERE VACA_TIPO = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboEstado.Clear
cboEstado.AddItem "Solicitado"
cboEstado.Text = "Solicitado"


rbAccion.Item(0).Value = True

If rs!REQUIERE_AUTORIZACION = 0 Then
    cboEstado.AddItem "Autorizado"
End If

If rs!PERMITE_LIQUIDACION = 1 Then
    rbAccion.Item(1).Enabled = True
Else
    rbAccion.Item(1).Enabled = False
End If

Call rbAccion_Click(0)

rs.Close
End Sub


Private Sub dtpFechaC_Change()
If txtEmpleadoId.Text <> "" Then
    txtDias.Text = fxRH_Dias_Laborales(txtEmpleadoId.Text, dtpFechaI.Value, dtpFechaC.Value)
End If
End Sub

Private Sub dtpFechaI_Change()
If txtEmpleadoId.Text <> "" Then
    txtDias.Text = fxRH_Dias_Laborales(txtEmpleadoId.Text, dtpFechaI.Value, dtpFechaC.Value)
End If
End Sub

Private Sub Form_Load()

vModulo = 23

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbInicializa()

On Error GoTo vError

vPaso = True

    dtpFechaI.Value = fxFechaServidor
    dtpFechaC.Value = dtpFechaI.Value

    strSQL = "select VACA_TIPO as Idx, rtrim(Descripcion) as ItmX" _
           & " from RH_VACACIONES_TIPOS" _
           & " Where Activa = 1"
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

vPaso = False

txtEmpleadoId.SetFocus


Call cboTipo_Click


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub rbAccion_Click(Index As Integer)

If Index = 0 Then
    dtpFechaI.Enabled = True
    dtpFechaC.Enabled = True
    
    txtDias.Locked = True
    
    Call dtpFechaC_Change
    
Else
    dtpFechaI.Enabled = False
    dtpFechaC.Enabled = False
    txtDias.Locked = False

End If

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
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
   txtIdentificacion.Text = Trim(gBusquedas.Resultado2)
   txtNombre.Text = gBusquedas.Resultado3
    
   Call sbConsulta
    
End Sub

Public Sub sbConsulta_Externa(pEmpleadoId As String)

txtEmpleadoId.Text = pEmpleadoId
Call sbConsulta

End Sub


Private Sub sbConsulta()


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vRH_Vacaciones_Info" _
       & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtDiasDisponibles.Text = Format(rs!Dias_Disponibles, "Standard")
    txtDias.Text = 1
    
    
    dtpFechaI.MinDate = rs!Fecha_Inicio
    dtpFechaC.MinDate = rs!Fecha_Inicio
    
    dtpFechaI.Value = rs!Fecha
    dtpFechaC.Value = rs!Fecha
    
    txtEmpleadoId.Text = rs!Empleado_ID
    txtIdentificacion.Text = rs!IDENTIFICACION
    txtNombre.Text = rs!Nombre
    
    mNomina = rs!COD_NOMINA
    
Else
    txtDiasDisponibles.Text = 0
    txtDias.Text = 0
End If
rs.Close

Call cboTipo_Click

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtEmpleadoId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub
