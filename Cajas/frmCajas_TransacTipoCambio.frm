VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCajas_TransacTipoCambio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compra / Venta de Divisas"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Compra"
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
      Value           =   -1  'True
   End
   Begin VB.Timer TimerCaja 
      Interval        =   10
      Left            =   10080
      Top             =   240
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5340
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
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
            TextSave        =   "10:17:p. m."
            Object.ToolTipText     =   "Fecha/Hora"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3831
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
      Height          =   315
      Left            =   4560
      TabIndex        =   8
      Top             =   240
      Width           =   5535
      _Version        =   1310723
      _ExtentX        =   9763
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
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   10215
      _Version        =   1310723
      _ExtentX        =   18018
      _ExtentY        =   2778
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   2772
         _Version        =   1310723
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   312
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   12
         Top             =   600
         Width           =   5412
         _Version        =   1310723
         _ExtentX        =   9546
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   795
         Index           =   0
         Left            =   7080
         TabIndex        =   13
         Top             =   600
         Width           =   855
         _Version        =   1310723
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Pago"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCajas_TransacTipoCambio.frx":0000
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   795
         Index           =   1
         Left            =   8040
         TabIndex        =   14
         Top             =   600
         Width           =   855
         _Version        =   1310723
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCajas_TransacTipoCambio.frx":0462
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   795
         Index           =   2
         Left            =   8880
         TabIndex        =   15
         Top             =   600
         Width           =   975
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Cancelar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCajas_TransacTipoCambio.frx":0C3A
         TextImageRelation=   1
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   20
      Top             =   1080
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Venta"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   3600
      TabIndex        =   21
      Top             =   2880
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
      Height          =   315
      Left            =   3600
      TabIndex        =   22
      Top             =   2040
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtImporte 
      Height          =   315
      Left            =   3600
      TabIndex        =   23
      Top             =   2400
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   556
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
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   330
      Left            =   3600
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monto de Cambio ...:"
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
      Index           =   3
      Left            =   1680
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Divisa  ...:"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio ...:"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa ...:"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblTransaccion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gesstión ...:"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmCajas_TransacTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mToken As String
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cboDivisa_Click()
If vPaso Then Exit Sub

If cboDivisa.ListCount = 0 Then Exit Sub





End Sub

Private Sub Form_Activate()
 vModulo = 5

End Sub

Private Sub Form_Load()
vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mToken = CStr(Hour(Time))

txtCedula.Text = ModuloCajas.mClienteId
txtNombre.Text = ModuloCajas.mCliente

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerCaja_Timer()

TimerCaja.Interval = 0
TimerCaja.Enabled = False

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
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

Me.Caption = "Cambio de Divisas       ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario
           

StatusBarX.Panels(1).Text = ModuloCajas.mDescripcion
StatusBarX.Panels(2).Text = ModuloCajas.mOficinaDesc

'Inicializa datos Principales
Call sbInicializa

End Sub



Private Sub sbInicializa()

ModuloCajas.mTiquete = ""

txtCedula.Text = ModuloCajas.mClienteId
txtNombre = ModuloCajas.mCliente

strSQL = "select rtrim(Doc.Tipo_Documento) as  'Idx', Doc.Descripcion as 'Itmx'" _
       & " from Cajas_Documentos Cj inner join SIF_Documentos Doc on Cj.Tipo_Documento = Doc.Tipo_Documento" _
       & " Where Cj.Cod_Caja = '" & ModuloCajas.mCaja & "'"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

vPaso = True
    strSQL = "select COD_DIVISA as 'IdX', DESCRIPCION as 'ItmX' " _
           & "  from vSys_Divisas Where DIVISA_LOCAL = 0"
    Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)
vPaso = False

End Sub
