VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_Excedentes_Pago 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excedentes: Auxiliar de Pago"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   8055
      _Version        =   1441793
      _ExtentX        =   14208
      _ExtentY        =   10610
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
      SelectedItem    =   1
      Item(0).Caption =   "Pagos"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "Label3(0)"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "Label3(2)"
      Item(0).Control(3)=   "Label3(3)"
      Item(0).Control(4)=   "Label3(4)"
      Item(0).Control(5)=   "rbProceso(0)"
      Item(0).Control(6)=   "rbProceso(1)"
      Item(0).Control(7)=   "rbProceso(2)"
      Item(0).Control(8)=   "rbProceso(3)"
      Item(0).Control(9)=   "rbProceso(4)"
      Item(0).Control(10)=   "rbProceso(5)"
      Item(0).Control(11)=   "gbSep1(0)"
      Item(1).Caption =   "Reportes"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbSep1(1)"
      Item(1).Control(1)=   "gbSep1(2)"
      Begin XtremeSuiteControls.GroupBox gbSep1 
         Height          =   855
         Index           =   0
         Left            =   -69640
         TabIndex        =   15
         Top             =   5040
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12938
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnAplicar 
            Height          =   495
            Left            =   5880
            TabIndex        =   16
            Top             =   240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmAH_Excedentes_Pago.frx":0000
            ImageAlignment  =   4
         End
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   0
         Left            =   -69040
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Separar Salidas para Pagos (Integral)"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   1
         Left            =   -69040
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asignacion de Casos Especiales"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   2
         Left            =   -69040
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Acreditar Cuentas de Ahorros"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   3
         Left            =   -69040
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Enviar a Tesorería"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   4
         Left            =   -69040
         TabIndex        =   13
         Top             =   4560
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Enviar a Fondos de ahorros"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   5
         Left            =   -69040
         TabIndex        =   14
         Top             =   3840
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Reclasificar Salidas"
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
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gbSep1 
         Height          =   855
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   4920
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12938
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnInforme 
            Height          =   495
            Left            =   5880
            TabIndex        =   18
            Top             =   240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Reporte"
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
            Picture         =   "frmAH_Excedentes_Pago.frx":0727
            ImageAlignment  =   4
         End
      End
      Begin XtremeSuiteControls.GroupBox gbSep1 
         Height          =   3375
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   7455
         _Version        =   1441793
         _ExtentX        =   13150
         _ExtentY        =   5953
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   20
            Top             =   720
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Resumen de Salidas"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   21
            Top             =   1200
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detalle de Excedentes"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   22
            Top             =   1680
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Dimex Inactivos"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   23
            Top             =   2160
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Más de una Cuenta (Ahorros Interna)"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   24
            Top             =   2640
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detalle de Pago de Excedentes"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   0
            Width           =   7335
            _Version        =   1441793
            _ExtentX        =   12938
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Informes de Pago de Excedentes"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   4
         Left            =   -69640
         TabIndex        =   8
         Top             =   4200
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 5: Traslado de Fondos de Ahorros"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   3
         Left            =   -69640
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 4: Traslado a Bancos (Tesorería)"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   2
         Left            =   -69640
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 3: Transferencias Internas"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   1
         Left            =   -69640
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 2: Casos Especiales"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   0
         Left            =   -69640
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 1: Separar Casos"
         ForeColor       =   16711680
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
      End
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Periodo"
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
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Auxiliar de Pagos de Excedentes"
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
      Height          =   492
      Index           =   11
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   9252
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmAH_Excedentes_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub btnInforme_Click()
Dim strSQL As String
Dim pCorte As Date, pCorteFiltro As String


On Error GoTo vError


Me.MousePointer = vbHourglass


With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Excedentes - Reportes"
    
    .Connect = glogon.ConectRPT
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "DD/MM/YYYY") & "'"
    .Formulas(3) = "usuario='" & UCase(glogon.Usuario) & "'"
    .Formulas(4) = "subtitulo='PERIODO: " & cboPeriodo.Text & "'"
    
    

    Select Case True
      Case rbInforme(0).Value 'Resumen de Salidas
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Salidas_Resumen.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
        
      Case rbInforme(1).Value 'Detalle
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_POS_CIERRE.rpt")
        .SelectionFormula = "{EXC_CIERRE.ID_PERIODO} = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
      
      Case rbInforme(2).Value 'Casos con Dimex inactivos
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Dimex_Inactivo_Lista.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario

      Case rbInforme(3).Value 'Personas con mas de una cuenta Sinpe
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Personas_ConMasCtasSinpe.rpt")
        .StoredProcParam(0) = glogon.Usuario

      Case rbInforme(4).Value 'Detalle de Pago
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Salidas_Pago_Detalle.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
    End Select

    .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Me.MousePointer = vbHourglass

vPaso = True


strSQL = "select IdX, ItmX from vExc_Periodos where estado in('C') order by Idx desc"
Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)

vPaso = False

Me.MousePointer = vbDefault

End Sub
