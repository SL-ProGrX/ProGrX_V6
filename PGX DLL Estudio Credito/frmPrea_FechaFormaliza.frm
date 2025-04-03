VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPrea_FechaFormaliza 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fecha Formalización"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.FlatEdit txtFechaCorte 
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1080
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtDias 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2280
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2640
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFormaliza 
      Height          =   330
      Left            =   3480
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
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
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton btnCalcular 
      Height          =   570
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Cambio de Evaluador"
      Top             =   3480
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   1005
      _StockProps     =   79
      Caption         =   "Calcular"
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
      Appearance      =   21
      Picture         =   "frmPrea_FechaFormaliza.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCambiar 
      Height          =   570
      Left            =   4080
      TabIndex        =   10
      ToolTipText     =   "Cambio de Evaluador"
      Top             =   3480
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   1005
      _StockProps     =   79
      Caption         =   "Cambiar"
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
      Appearance      =   21
      Picture         =   "frmPrea_FechaFormaliza.frx":06E5
      ImageAlignment  =   4
   End
   Begin XtremeShortcutBar.ShortcutCaption scFechas 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   6255
      _Version        =   1572864
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Interés calculado"
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
      Height          =   330
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Formalización"
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
      Height          =   210
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dias Interés"
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
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Corte"
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
      Height          =   210
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   615
      Index           =   9
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16960
      _ExtentY        =   1085
      _StockProps     =   14
      Caption         =   "Calcular Intereses Proyectando la Fecha de Formalización"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmPrea_FechaFormaliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mMonto As Currency, mTasa As Currency

Private Sub btnCalcular_Click()

On Error GoTo vError


strSQL = "exec spCrdPrea_InteresesFormaliza_Calculo '" & GLOBALES.gTag & "', " & mMonto & ", " & mTasa _
       & ", '" & Format(dtpFormaliza.Value, "yyyy-mm-dd") & "', '" & txtFechaCorte.Text & "'"
Call OpenRecordSet(rs, strSQL)

txtDias.Text = CStr(rs!Dias)
txtMonto.Text = Format(rs!Monto_Interes, "Standard")


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub btnCambiar_Click()

On Error GoTo vError

strSQL = "exec spCrdPreaGuardaMontoInteresProyectado '" & GLOBALES.gTag & "', " & CCur(txtMonto.Text) & ", 1"
Call ConectionExecute(strSQL)

UnLoad Me

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub Form_Load()

On Error GoTo vError

'fxCRDPreaCalcularGastosCredito](@Monto dec(14,2), @Tasa dec(10,2), @DiasInteres dec(10,2), @Linea varchar(10), @Garantia char(1)) returns dec(10,2)
'fxCrdPreaCalculaInteresPreanalisis]  (@MontoCredito DECIMAL(18,2),@TasaInteres DECIMAL(18,2),@FecProyeccionInt DATE, @Cedula VARCHAR(20))
'fxCrdPreaConsultaComisionPorLinea]
'           (@Linea varchar(10),@CodDestino varchar(5),@Garantia char(3))


strSQL = "exec spCrdPrea_FechaCalIntereses '" & GLOBALES.gTag & "'"
Call OpenRecordSet(rs, strSQL)

mMonto = rs!Monto
mTasa = rs!TASA

scFechas.Caption = "Planilla Recibida: " & Format(rs!Planilla_Aplica, "yyyy-mm-dd") & "   ¦  Planilla Enviada: " & Format(rs!Planilla_Envio, "yyyy-mm-dd")

txtFechaCorte.Text = Format(rs!Fecha_Corte, "yyyy-mm-dd")

dtpFormaliza.MaxDate = rs!Fecha_Corte
dtpFormaliza.Value = rs!Formalizacion
dtpFormaliza.MinDate = rs!Formalizacion

txtDias.Text = 0
txtMonto.Text = Format(0, "Standard")

Exit Sub

vError:

End Sub
