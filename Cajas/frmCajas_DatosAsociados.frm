VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCajas_DatosAsociados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Transacciones "
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   12450
   Begin XtremeSuiteControls.PushButton btnIncobrables 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   6360
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Incobrables"
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
      Picture         =   "frmCajas_DatosAsociados.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.CheckBox chkSF_Liquidados 
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   5890
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Saldos a Favor Liquidados/Cancelados"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   0
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Pago: Conceptos de Cajas"
      Top             =   1080
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":0720
   End
   Begin VB.Timer TimerCaja 
      Interval        =   10
      Left            =   360
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6810
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Caja"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            TextSave        =   "22:40"
            Object.ToolTipText     =   "Fecha/Hora"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Saldo a Favor"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraCajaInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Label lblInfoUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6120
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblInfoApertura 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblInfoCaja 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario en uso ..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Apertura ..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Caja ..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5055
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Presione el botón para realizar el tramite"
      Top             =   1080
      Width           =   11295
      _Version        =   524288
      _ExtentX        =   19923
      _ExtentY        =   8916
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
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmCajas_DatosAsociados.frx":10C6
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboBusqueda 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   582
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
      Height          =   330
      Left            =   5040
      TabIndex        =   11
      Top             =   480
      Width           =   5535
      _Version        =   1572864
      _ExtentX        =   9763
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   1
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Retiros: Caja Chica"
      Top             =   1680
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":32C7
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   2
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Abonos a Créditos"
      Top             =   2400
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":3C2B
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   3
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Aportes de Fondos"
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":4576
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   4
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Abonos a Cuentas por Cobrar"
      Top             =   3600
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":4F36
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   5
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Intercambio de Valores"
      Top             =   4320
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":5994
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   612
      Index           =   6
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Saldos a favor"
      Top             =   5040
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      Transparent     =   -1  'True
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_DatosAsociados.frx":62D4
   End
   Begin XtremeSuiteControls.PushButton btnPatrimonio 
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   6360
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Patrimonio"
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
      Picture         =   "frmCajas_DatosAsociados.frx":6C38
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnSesion 
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   22
      Top             =   6360
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ver Sesión activa"
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
      Picture         =   "frmCajas_DatosAsociados.frx":7358
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnSesion 
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   23
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Finalizar Sesión"
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
      Picture         =   "frmCajas_DatosAsociados.frx":7A58
      ImageAlignment  =   0
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCajas_DatosAsociados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRA_Access As Boolean

Private Sub sbConsultaCreditos()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "exec spCajas_Consulta_Creditos '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

vGrid.ActiveSheet = 1
vGrid.Sheet = 1
vGrid.MaxRows = 0

Do While Not rs.EOF
 vGrid.MaxRows = vGrid.MaxRows + 1
 vGrid.Row = vGrid.MaxRows
 vGrid.Col = 2
 vGrid.Text = Trim(rs!ID_SOLICITUD)
 vGrid.Col = 3
 vGrid.Text = Trim(rs!Codigo)
 vGrid.Col = 4
 vGrid.Text = Trim(rs!GarantiaDesc)
 vGrid.Col = 5
 vGrid.Text = Format(rs!Saldo, "Standard")
 vGrid.Col = 6
 vGrid.Text = Format(rs!Mora, "Standard")
 vGrid.Col = 7
 vGrid.Text = Format(rs!Cuota, "Standard")
 vGrid.Col = 8
 vGrid.Text = Trim(rs!LineaDesc)
 
 rs.MoveNext
Loop
rs.Close

End Sub


Private Sub btnAcciones_Click(Index As Integer)
On Error GoTo vError


ModuloCajas.mClienteId = txtCedula.Text
ModuloCajas.mCliente = txtNombre.Text
    
If chkSF_Liquidados.Visible Then
    chkSF_Liquidados.Visible = False
'    Me.Height = Me.Height - 330
End If

Select Case Index
  Case 0  'General
    Call sbConsultaServicios
    Call sbFormsCall("frmCajas_Transacciones", vbModal, 1, 1, False, Me)

  Case 1 'Caja Chica
    Call sbConsultaServicios
    Call sbFormsCall("frmCajas_CajaChica", vbModal, 1, 1, False, Me)
  
  Case 2 'Creditos"
    Call sbConsultaCreditos
  
  Case 3 'Fondos
    Call sbConsultaFondos
  
  Case 4 'CxC
    Call sbConsultaCuentas
  
  Case 5 'Cambio
    Call sbFormsCall("frmCajas_TransacTipoCambio", vbModal, 1, 1, False, Me)
  
  Case 6 'Saldos
    chkSF_Liquidados.Visible = True
    Me.Height = Me.Height + 330
    
    Call sbConsultaSaldosAfavor
    
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub btnIncobrables_Click()
MsgBox "Esta opción está bajo revisión para Cambio Estructural!", vbExclamation
End Sub

Private Sub btnPatrimonio_Click()
On Error GoTo vError

If txtCedula.Text = "" Or txtNombre.Text = "" Then Exit Sub

GLOBALES.gCedulaActual = txtCedula.Text
frmAH_RegistraAhorro.Show vbModal

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnSesion_Click(Index As Integer)
Call sbFormsCall("frmCajas_Sesion", vbModal, , , False, Me, True)
End Sub

Private Sub chkSF_Liquidados_Click()
Call sbConsultaSaldosAfavor

End Sub

Private Sub Form_Activate()
vModulo = 5

End Sub

Private Sub Form_Load()

vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

vGrid.MaxRows = 0

cboBusqueda.Clear
cboBusqueda.AddItem "Identificación"
cboBusqueda.AddItem "No. Operación"

cboBusqueda.Text = "Identificación"

Call RefrescaTags(Me)
Call Formularios(Me)

End Sub





Private Sub TimerCaja_Timer()

TimerCaja.Interval = 0
TimerCaja.Enabled = False


ModuloCajas.mClienteId = ""
ModuloCajas.mCliente = ""

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

'Paso 3: Continuar con Barra de Información
'lblInfoApertura.Caption = ModuloCajas.mApertura
'lblInfoCaja.Caption = ModuloCajas.mCaja
'lblInfoUsuario.Caption = ModuloCajas.mUsuario

Me.Caption = "Transacciones de Cajas      ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

StatusBarX.Panels(1).Text = ModuloCajas.mDescripcion
StatusBarX.Panels(2).Text = ModuloCajas.mOficinaDesc


End Sub


Private Sub sbConsultaCuentas(Optional pSheet As Integer = 3)
Dim strSQL As String, rs As New ADODB.Recordset

Dim curCuota As Currency, curMonto As Currency
Dim curSaldo As Currency, vMora As Boolean
Dim i As Integer

On Error GoTo vError

curCuota = 0
curMonto = 0
curSaldo = 0

vMora = False


'txtTotalMonto.Text = ""
'txtTotalSaldo.Text = ""
'txtTotalCuota.Text = ""
'
'StatusBar.Panels(5).Text = "0.00"
'StatusBar.Panels(6).Text = "0.00"

Me.MousePointer = vbHourglass


vMora = False

With vGrid
 .Sheet = pSheet
 .ActiveSheet = pSheet
 
 
 .MaxRows = 0
 strSQL = "exec spCxC_PersonasCuentas '" & txtCedula.Text & "','A'"
 
 rs.CursorLocation = adUseServer
 Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    
    For i = 2 To .MaxCols
      .Col = i
      Select Case i
        Case 2 'Operacion

           .Text = CStr(rs!Operacion)

'              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        
       
            If rs!Warning = 1 Then
'               .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
              .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
              .CellNoteIndicatorColor = vbRed
              .CellNote = "Dias para el vencimiento : " & DateDiff("d", rs!Fecha_Server, rs!Fecha_Pago)
            End If
        
             
             If Mid(rs!Estado, 1, 1) = "C" Then
'                .TypePictPicture = imgSemaforos.ListImages.Item(6).Picture
             End If

            'Indicador de Morosidad
            If rs!MoraMonto > 0 Then
              
'              .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
              vMora = True
            
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
            
              .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
              .CellNoteIndicatorColor = vbBlue
              
              .CellNote = "Morosidad:" & vbCrLf _
                        & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                        & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                        & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                        & "   Días Mora : " & Format(rs!MoraDias, "###,##0") & vbCrLf _
                        & "   Cta. Ultima : " & Format(rs!MoraFecha, "dd-mm-yyyy") & vbCrLf & vbCrLf _
                        & "   Total Mora  : " & Format(rs!MoraMonto, "Standard") & vbCrLf
            
            End If
        
        
        Case 3 'Concepto
            .Text = rs!cod_Concepto
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = Trim(rs!ConceptoDesc) & vbCrLf & vbCrLf & "Activación: " & Format(rs!Activa_Fecha, "dd/mm/yyyy") & vbCrLf & "Usuario: " & Trim(rs!ACTIVA_USUARIO) & vbCrLf & "Oficina:" & rs!OficinaDesc & ""
        
        Case 4 'Documento
            .Text = Trim(rs!Num_Documento & "")
        Case 5 'Fecha Pago
            .Text = Format(rs!Fecha_Pago, "dd/mm/yyyy")
        Case 6 'Monto
            .Text = Format(rs!Monto, "Standard")
        Case 7 'Saldo
            .Text = Format(rs!Saldo, "Standard")
        Case 8 'Cuota
            .Text = Format(rs!Cuota, "Standard")
        Case 9 'Mora Dias
            .Text = CStr(rs!MoraDias)
        
        Case 10 'Estado
            .Text = CStr(rs!Estado)
            
      
      End Select
    Next i
    
     curMonto = curMonto + rs!Monto
     curSaldo = curSaldo + rs!Saldo
     curCuota = curCuota + rs!Cuota

    rs.MoveNext
  Loop
  rs.Close
  
End With

  
'Totales
'txtTotalMonto.Text = Format(curMonto, "Standard")
'txtTotalCuota.Text = Format(curCuota, "Standard")
'txtTotalSaldo.Text = Format(curSaldo, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Function fxNumeroCedula() As String
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxNumeroCedula = ""

strSQL = "select cedula from reg_Creditos where id_solicitud = " & txtCedula.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   fxNumeroCedula = Trim(rs!Cedula & "")
End If
rs.Close

Exit Function

vError:
    fxNumeroCedula = ""

End Function

Private Function fxPersonaSaldoFavor() As Currency
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxPersonaSaldoFavor = 0

strSQL = "select sum(Saldo) as Saldo" _
       & " From CAJAS_SALDO_FAVOR" _
       & " where cedula = '" & Trim(txtCedula.Text) & "'"
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   fxPersonaSaldoFavor = rs!Saldo
End If
rs.Close

Exit Function

vError:
    fxPersonaSaldoFavor = 0

End Function



Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtCedula.Text) <> "" Then
    
    If cboBusqueda.Text <> "Identificación" Then
       txtCedula.Text = fxNumeroCedula
       cboBusqueda.Text = "Identificación"
    End If
    

    txtNombre.Text = fxNombre(txtCedula.Text)
    Call sbConsultaCreditos
End If

If KeyCode = vbKeyF4 Then
  If cboBusqueda.Text = "Identificación" Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id. Alterno"
    gBusquedas.Col3Name = "Nombre"
    
    gBusquedas.Consulta = "select cedula,cedular,nombre from socios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = ""
    gBusquedas.Convertir = "N"
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado3
   Else
   
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id. Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "select S.cedula,S.cedular, S.Nombre,R.id_solicitud as 'No.Operación'" _
                        & " from socios S inner join reg_Creditos R on S.cedula = R.cedula"
    gBusquedas.Columna = "R.id_solicitud"
    gBusquedas.Orden = "R.id_solicitud"
    gBusquedas.Filtro = ""
    gBusquedas.Convertir = "N"
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado3
   
   End If

End If

End Sub

Private Sub txtCedula_LostFocus()
    'Valida Acceso a Expediente
    vRA_Access = fxSys_RA_Consulta(Trim(txtCedula.Text), glogon.Usuario)
     
    If Not vRA_Access Then
        MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
        txtCedula.Text = ""
        txtNombre.Text = ""
        Exit Sub
    End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    
    gBusquedas.Consulta = "select cedula,nombre from socios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = ""
    gBusquedas.Convertir = "N"
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
End If
End Sub



Private Sub sbConsultaFondos()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_Consulta_Fondos '" & Trim(txtCedula.Text) & "','" & glogon.Usuario & "'"


Call OpenRecordSet(rs, strSQL)
vGrid.ActiveSheet = 2
vGrid.Sheet = 2
vGrid.MaxRows = 0


  Do While Not rs.EOF
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    vGrid.Col = 2
    vGrid.Text = CStr(rs!COD_CONTRATO)
    vGrid.Col = 3
    vGrid.Text = Trim(rs!COD_PLAN)
    vGrid.Col = 4
    vGrid.Text = Format(rs!Aportes, "Standard")
    vGrid.Col = 5
    vGrid.Text = Format(rs!Rendimiento, "Standard")
    vGrid.Col = 6
    vGrid.Text = Format(rs!acumulado, "Standard")
    vGrid.Col = 7
    vGrid.Text = Format(rs!Monto, "Standard")
    vGrid.Col = 8
    vGrid.Text = Trim(rs!PlanDesc)
    
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsultaServicios()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_Consulta_Servicios '" & Trim(txtCedula.Text) & "'"
Call OpenRecordSet(rs, strSQL)
vGrid.ActiveSheet = 4
vGrid.Sheet = 4
vGrid.MaxRows = 0


  Do While Not rs.EOF
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
         
    For i = 1 To vGrid.MaxCols
      vGrid.Col = i
      Select Case i
         Case 1 'Servicio
            vGrid.Text = rs!ServicioDesc
         Case 2 'Monto
            vGrid.Text = Format(rs!Monto, "Standard")
         Case 3 'Fecha
            vGrid.Text = Format(rs!Monto, "Standard")
         Case 4 'No.Ref.
            vGrid.Text = rs!Num_Referencia
         Case 5 'Caja
            vGrid.Text = rs!COD_CAJA
         Case 6 'Usuario
            vGrid.Text = rs!REGISTRO_USUARIO
         Case 7 'Recaudador
            vGrid.Text = rs!RecaudadorDesc
         
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsultaSaldosAfavor()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *" _
       & " From CAJAS_SALDO_FAVOR" _
       & " where cedula = '" & Trim(txtCedula.Text) & "'"

If chkSF_Liquidados.Value = vbChecked Then
    strSQL = strSQL & " and Saldo <= 0"
Else
    strSQL = strSQL & " and Saldo > 0"
End If

strSQL = strSQL & " Order by Registro_Fecha desc"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
vGrid.ActiveSheet = 5
vGrid.Sheet = 5
vGrid.MaxRows = 0


  Do While Not rs.EOF
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
         
    For i = 3 To vGrid.MaxCols
      vGrid.Col = i
      Select Case i
         Case 3 'Linea
            vGrid.Text = CStr(rs!Linea)
         Case 4 'Tipo Doc
            vGrid.Text = rs!DOC_TIPO
         Case 5 'No.Doc.
            vGrid.Text = rs!Doc_Numero
         Case 6 'Registro Fecha
            vGrid.Text = rs!REGISTRO_FECHA
         Case 7 'Monto
            vGrid.Text = Format(rs!Monto, "Standard")
         Case 8 'Saldo
            vGrid.Text = Format(rs!Saldo, "Standard")
         Case 9 'Referencias
            vGrid.Text = "Tes. Id.: " & rs!Doc_Transac_Id & " ¦ Caja .: " & rs!COD_CAJA & "  Ap.Id.: " & rs!Cod_Apertura
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbConsultaRecibosMultiples()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *" _
       & " From CAJAS_AM_MAIN" _
       & " where cedula = '" & Trim(txtCedula.Text) & "'"
strSQL = strSQL & " Order by Registro_Fecha desc"

Call OpenRecordSet(rs, strSQL)
vGrid.ActiveSheet = 6
vGrid.Sheet = 6
vGrid.MaxRows = 0


  Do While Not rs.EOF
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows

    For i = 2 To vGrid.MaxCols
      vGrid.Col = i
      Select Case i
         Case 2 'Recibo Id
            vGrid.Text = CStr(rs!CAJA_AM_ID)
         Case 3 'Monto
            vGrid.Text = Format(rs!Monto, "Standard")
         Case 4 'Fecha
            vGrid.Text = rs!REGISTRO_FECHA
         Case 5 'Caja
            vGrid.Text = rs!COD_CAJA
         Case 6 'Apertura
            vGrid.Text = rs!Cod_Apertura
         Case 7 'Usuario
            vGrid.Text = rs!REGISTRO_USUARIO
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form

'vGrid.MaxRows = 0

If Not vRA_Access Then Exit Sub

Select Case vGrid.ActiveSheet
 Case 1 'Creditos
    vGrid.ActiveSheet = 1
    vGrid.Sheet = 1
    
    vGrid.Row = Row
    vGrid.Col = 2
    
    ModuloCajas.mRef_01 = vGrid.Text
    If GLOBALES.SysPlanPagos = 1 Then
        Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, , , False, Me, True)
    Else
        Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, , , False, Me, True)
    End If
    Call sbConsultaCreditos
 
 Case 2 'Fondos
    vGrid.ActiveSheet = 2
    vGrid.Sheet = 2
    vGrid.Row = Row
    vGrid.Col = 3
    gFondos.Plan = vGrid.Text
    vGrid.Col = 2
    gFondos.Contrato = vGrid.Text
    gFondos.Operadora = 1
    Call sbFormsCall("frmCajas_FNDAportaciones", vbModal, , , False, Me)
    Call sbConsultaFondos
    
 Case 3 'CxC / Factoreo
    vGrid.ActiveSheet = 3
    vGrid.Sheet = 3
    
    vGrid.Row = Row
    vGrid.Col = 2
    
    ModuloCajas.mRef_01 = vGrid.Text
    Call sbFormsCall("frmCxC_CuentasAbonos", vbModal, , , False, Me)
    Call sbConsultaCuentas

 Case 5 'Saldos a Favor
    ModuloCajas.mClienteId = txtCedula.Text
    ModuloCajas.mCliente = txtNombre.Text
    If Col = 1 Then
        Call sbFormsCall("frmCajas_TransacSFLiq", vbModal, 1, 1, False, Me)
        Call sbConsultaSaldosAfavor
    End If
    
    'Comprobante de Aplicación
    If Col = 2 Then
       vGrid.Row = Row
       vGrid.Col = 3
       
       Call sbComprobanteSF(vGrid.Text)
    End If



 Case 6 'Imprime Recibo Multiple
    If Col = 1 Then
       vGrid.Row = Row
       vGrid.Col = 2
       
       Call sbCaja_Recibo_Multiple(vGrid.Text)
    End If


End Select

'Muestar el Saldo a Favor
StatusBarX.Panels(4).Text = "Saldo Favor: " & Format(fxPersonaSaldoFavor, "Standard")

End Sub

Private Sub sbComprobanteSF(pId As Long)
Dim strSQL As String, x As New clsImpresoras
Dim vFlat As Boolean, rs As New ADODB.Recordset
Dim vEmpresa As String, vCedJur As String
Dim vArchivo As String

On Error GoTo vError

strSQL = "select nombre,cedula_juridica from sif_empresa"
Call OpenRecordSet(rs, strSQL)
 vEmpresa = UCase(rs!Nombre & "")
 vCedJur = Trim(rs!cedula_juridica & "")
rs.Close

With frmContenedor.Crt
   .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Cajas: Comprobante de Descargo de Saldos a Favor"
   
   .Connect = glogon.ConectRPT
   
    x.TipoImpresora = Recibos
    x.Reset
    .PrinterDriver = x.Controlador
    .PrinterName = x.Nombre
    .PrinterPort = x.Puerto
    
    .PrinterSelect
    
    .Destination = crptToWindow
     
    .Formulas(0) = "fxEmpresa = '" & vEmpresa & "'"
    .Formulas(1) = "fxCedJur = '" & vCedJur & "'"
    .Formulas(2) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(3) = "fxFecha = '" & fxFechaServidor & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Cajas_SF_Comprobante.rpt")
    
    .SelectionFormula = "{CAJAS_SALDO_FAVOR.LINEA} = " & pId
    
   .PrintReport
End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)

ModuloCajas.mClienteId = txtCedula.Text
ModuloCajas.mCliente = txtNombre.Text
    
If chkSF_Liquidados.Visible Then
    chkSF_Liquidados.Visible = False
'    Me.Height = Me.Height - 330
End If

vGrid.ActiveSheet = NewSheet

Select Case NewSheet
  Case 1 'Creditos
     Call sbConsultaCreditos
  Case 2 'Fondos
     Call sbConsultaFondos
  Case 3 'Cuentas x Cobrar
     Call sbConsultaCuentas
  Case 4 'Historico de Servicios
     Call sbConsultaServicios
     
  Case 5 'Saldos a Favor
    chkSF_Liquidados.Visible = True
'    Me.Height = Me.Height + 330
     
     Call sbConsultaSaldosAfavor

   Case 6 'Recibos Multiples
     Call sbConsultaRecibosMultiples
   
End Select

End Sub
