VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmRH_Permisos_Registro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Registro de Permisos"
   ClientHeight    =   4680
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9495
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8760
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   9495
      _Version        =   1441792
      _ExtentX        =   16748
      _ExtentY        =   5530
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   612
         Left            =   7200
         TabIndex        =   1
         Top             =   2520
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmRH_Permisos_Registro.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   6852
         _Version        =   1441792
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
         _Version        =   1441792
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
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
         _Version        =   1441792
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
      Begin XtremeSuiteControls.DateTimePicker dtpHoraI 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
         _Version        =   1441792
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
         Format          =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHoraC 
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
         _Version        =   1441792
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
         Format          =   2
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   6600
         TabIndex        =   17
         Top             =   1680
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.FlatEdit txtHoras 
         Height          =   315
         Left            =   2160
         TabIndex        =   19
         ToolTipText     =   "Dias a Disfrutar"
         Top             =   2640
         Width           =   1335
         _Version        =   1441792
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
      Begin XtremeSuiteControls.FlatEdit txtHrsMax 
         Height          =   315
         Left            =   3480
         TabIndex        =   21
         ToolTipText     =   "Dias a Disfrutar"
         Top             =   2640
         Width           =   1335
         _Version        =   1441792
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
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
         Index           =   5
         Left            =   1080
         TabIndex        =   20
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
         Left            =   5520
         TabIndex        =   18
         Top             =   1680
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Horas"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   13
         Top             =   2640
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
         TabIndex        =   12
         Top             =   1680
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
         TabIndex        =   5
         Top             =   240
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
         TabIndex        =   4
         Top             =   600
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   1812
      _Version        =   1441792
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   5052
      _Version        =   1441792
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
      Left            =   480
      TabIndex        =   11
      Top             =   600
      Width           =   1812
      _Version        =   1441792
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
      Left            =   2280
      TabIndex        =   10
      Top             =   360
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
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   1692
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
      Left            =   4080
      TabIndex        =   8
      Top             =   360
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
Attribute VB_Name = "frmRH_Permisos_Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnAplicar_Click()
If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset
Dim Boleta As String, LiquidaId As Integer

'Validacion
If dtpHoraI.Value > dtpHoraC.Value Then
    MsgBox "Error en Rango de Horas!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtHoras.Text) Then
    MsgBox "Horas de Permisos Inválidas!", vbExclamation
    Exit Sub
End If

If CCur(txtHoras.Text) < 0 Then
    MsgBox "Horas de Permisos Inválidas!", vbExclamation
    Exit Sub
End If

If CCur(txtHoras.Text) > CCur(txtHrsMax.Text) Then
    MsgBox "Horas de Permisos Excedente el Total Permitido!", vbExclamation
    Exit Sub
End If

On Error GoTo vError

'spRH_Permisos_Registro(@EmpleadoId varchar(20), @Tipo varchar(10), @Notas varchar(1000), @Usuario varchar(30)
'                , @Salida datetime, @Entrada datetime, @Horas int, @PermisoFecha datetime
'                , @Estado char(1) = 'S', @AutorizaId varchar(30) = Null
'                , @AppCod varchar(30) = 'ProGrX' )
Dim pAutorizador As String

If Mid(cboEstado.Text, 1, 1) = "S" Then
  pAutorizador = "Null"
Else
  pAutorizador = "Null"
End If

strSQL = "exec spRH_Permisos_Registro '" & txtEmpleadoId.Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) _
        & "','" & txtNotas.Text & "','" & glogon.Usuario & "'" _
        & ",'" & Format(dtpFecha.Value, "yyyy/mm/dd") & " " & Format(dtpHoraI.Value, "hh:mm:ss") & "'" _
        & ",'" & Format(dtpFecha.Value, "yyyy/mm/dd") & " " & Format(dtpHoraC.Value, "hh:mm:ss") & "'" _
        & "," & CCur(txtHoras.Text) & ",'" & Format(dtpFecha.Value, "yyyy/mm/dd") _
        & "','" & Mid(cboEstado.Text, 1, 1) & "'," & pAutorizador & ",'ProGrX'"
Call OpenRecordSet(rs, strSQL)
    Boleta = rs!BoletaId
rs.Close

'Print Boleta
Call sbBoleta_Permisos(Boleta)

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

strSQL = "select REQUIERE_AUTORIZACION,PERMISO_HRS_MAX" _
       & " from RH_PERMISOS_TIPOS " _
       & " WHERE PERMISO_TIPO = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboEstado.Clear
cboEstado.AddItem "Solicitado"
cboEstado.Text = "Solicitado"

txtHrsMax.Text = CStr(rs!PERMISO_HRS_MAX)


If rs!REQUIERE_AUTORIZACION = 0 Then
    cboEstado.AddItem "Autorizado"
End If

rs.Close
End Sub

Private Sub Form_Load()

vModulo = 23

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vPaso = True
dtpFecha.Value = fxFechaServidor

dtpHoraI.Value = dtpFecha.Value
dtpHoraC.Value = dtpFecha.Value

'Tipo
  strSQL = "select Permiso_Tipo as Idx, rtrim(Descripcion) as ItmX from RH_PERMISOS_TIPOS"
  Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

'Estado
cboEstado.AddItem "Solicitado"
cboEstado.AddItem "Autorizado"
cboEstado.Text = "Solicitado"

vPaso = False

txtEmpleadoId.SetFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
Dim strSQL As String, rs As New Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select EMPLEADO_ID,IDENTIFICACION,NOMBRE_COMPLETO, dbo.Mygetdate() as 'Fecha' from Rh_Personas" _
       & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    dtpHoraI.Value = rs!Fecha
    dtpHoraC.Value = rs!Fecha
    dtpFecha.Value = rs!Fecha
    
    txtEmpleadoId.Text = rs!Empleado_ID
    txtIdentificacion.Text = rs!IDENTIFICACION
    txtNombre.Text = rs!NOMBRE_COMPLETO

    txtHoras.Text = 1
Else
    txtHoras.Text = 1
End If
rs.Close


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


