VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_ConInformeEspecial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe Especial de Consolidación"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkNotas_Patrimonio 
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   15
      Top             =   2520
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auxiliar de Patrimonio"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   13
      Top             =   1320
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Informe UENS Agrupado"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   6480
      Width           =   8055
      _Version        =   1572864
      _ExtentX        =   14208
      _ExtentY        =   2143
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   5880
         TabIndex        =   1
         Top             =   360
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_ConInformeEspecial.frx":0000
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   4920
      Width           =   8055
      _Version        =   1572864
      _ExtentX        =   14208
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Periodos:"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1296
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
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   312
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
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
      Begin XtremeSuiteControls.FlatEdit txtPeriodo 
         Height          =   312
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAnioCorte 
         Height          =   312
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   732
         _Version        =   1572864
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtMesCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
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
      Begin XtremeSuiteControls.FlatEdit txtPeriodoCorte 
         Height          =   312
         Left            =   2880
         TabIndex        =   9
         Top             =   720
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
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
         Transparent     =   -1  'True
      End
      Begin VB.Label Label3 
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
         Height          =   312
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   948
      End
      Begin VB.Label Label1 
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
         Height          =   312
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   948
      End
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   14
      Top             =   1920
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Notas a los Estados"
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
   End
   Begin XtremeSuiteControls.CheckBox chkNotas_Patrimonio 
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   16
      Top             =   2880
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auxiliar de Planes de Ahorros"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkNotas_Patrimonio 
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   17
      Top             =   3240
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auxiliar de Crédito"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkNotas_Patrimonio 
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   18
      Top             =   3600
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auxiliar de Activos Fijos"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkNotas_Patrimonio 
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   19
      Top             =   3960
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auxiliar de Cuentas por Pagar"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkNotas_Patrimonio 
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   20
      Top             =   4320
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auxiliar de Cuentas por Cobrar"
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
      Appearance      =   21
      Value           =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe de Consolidación"
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
      Height          =   615
      Index           =   1
      Left            =   1875
      TabIndex        =   12
      Top             =   360
      Width           =   5895
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmCntX_ConInformeEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Dim mContabilidad As Long, mAnio As Long, mMes As Integer

Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

mContabilidad = gCntX_Parametros.CodigoConta
mAnio = gCntX_Parametros.PeriodoAnio
mMes = gCntX_Parametros.PeriodoMes

txtAnio.Text = mAnio
txtMes.Text = mMes

txtAnioCorte.Text = mAnio
txtMesCorte.Text = mMes

txtPeriodo.Text = fxCntX_PeriodoDesc(mAnio, mMes)
txtPeriodoCorte.Text = fxCntX_PeriodoDesc(mAnio, mMes)


End Sub
