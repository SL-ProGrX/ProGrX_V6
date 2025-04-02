VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_CD_Liquidaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidaciones de Desembolsos por Cómites Sedes"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14895
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Datos"
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
      Picture         =   "frmAF_CD_Liquidaciones.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigoComite 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtRate 
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   661
      _StockProps     =   77
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cuentas"
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
      Picture         =   "frmAF_CD_Liquidaciones.frx":0719
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Reportes"
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
      Picture         =   "frmAF_CD_Liquidaciones.frx":0E21
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcionComite 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   7935
      _Version        =   1441793
      _ExtentX        =   13996
      _ExtentY        =   661
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   14895
      _Version        =   1441793
      _ExtentX        =   26273
      _ExtentY        =   12303
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
      ItemCount       =   3
      Item(0).Caption =   "Desembolsos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Detallar Liquidación"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "vGridOpxDetallar"
      Item(1).Control(1)=   "ShortcutCaption2(0)"
      Item(1).Control(2)=   "vGridFacturas"
      Item(1).Control(3)=   "ShortcutCaption2(1)"
      Item(1).Control(4)=   "GroupBox1"
      Item(2).Caption =   "Histórico"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGridHistorico"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2175
         Left            =   -65440
         TabIndex        =   15
         Top             =   4680
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   3836
         _StockProps     =   79
         Caption         =   "Detalle de la Liquidación"
         ForeColor       =   12582912
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
         Begin XtremeSuiteControls.FlatEdit txtTotal 
            Height          =   330
            Left            =   7800
            TabIndex        =   19
            Top             =   480
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
            _StockProps     =   77
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTotalFactura 
            Height          =   330
            Left            =   7800
            TabIndex        =   20
            Top             =   840
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
            _StockProps     =   77
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDiferencia 
            Height          =   330
            Left            =   7800
            TabIndex        =   21
            Top             =   1200
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
            _StockProps     =   77
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnLiquidar 
            Height          =   375
            Left            =   7800
            TabIndex        =   22
            Top             =   1680
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   661
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
            Picture         =   "frmAF_CD_Liquidaciones.frx":1528
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   1335
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   5535
            _Version        =   1441793
            _ExtentX        =   9763
            _ExtentY        =   2355
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notas de la Liquidación"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   10335
            _Version        =   1441793
            _ExtentX        =   18230
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Detalle de la Liquidación"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   18
            Top             =   1200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Diferencia"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   17
            Top             =   840
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto en Documentos"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   0
            Left            =   5880
            TabIndex        =   16
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total a Liquidar"
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6495
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   14655
         _Version        =   524288
         _ExtentX        =   25850
         _ExtentY        =   11456
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
         MaxCols         =   10
         SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":1C4F
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridHistorico 
         Height          =   6495
         Left            =   -70000
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   14895
         _Version        =   524288
         _ExtentX        =   26273
         _ExtentY        =   11456
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
         MaxCols         =   15
         SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":24D4
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridOpxDetallar 
         Height          =   5700
         Left            =   -69880
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   10054
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":353C
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridFacturas 
         Height          =   3615
         Left            =   -65440
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
         _ExtentY        =   6376
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_CD_Liquidaciones.frx":3B89
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Index           =   1
         Left            =   -65560
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Detalle de Facturas"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1441793
         _ExtentX        =   7646
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Cuentas Pendientes"
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
      End
   End
   Begin XtremeSuiteControls.PushButton btnDetallarLiquidacion 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Detallar la Liquidación"
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
      Picture         =   "frmAF_CD_Liquidaciones.frx":4225
      BorderGap       =   0
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15135
      _Version        =   1441793
      _ExtentX        =   26696
      _ExtentY        =   1508
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmAF_CD_Liquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim vOperacion As String, vDocumento As String, vDeposito As String, vDetalle As String, vFecha As String
Dim vMonto As Double

Private Sub btnBarra_Click(Index As Integer)
On Error GoTo vError

GLOBALES.gTag = txtCodigoComite.Text
Select Case Index
  Case 0 '"Datos"
    Call sbFormsCall("frmAF_CD_Comites", , , , False, Me)
        
  Case 1 '"Cuentas"
    Call sbFormsCall("frmAF_CD_Cuentas", , , , False, Me)
    
  Case 2 '"Reportes"
      strSQL = ""
      With frmContenedor.Crt
         .Reset
         .WindowShowGroupTree = True
         .WindowShowPrintSetupBtn = True
         .WindowShowRefreshBtn = True
         .WindowShowSearchBtn = True
         .WindowState = crptMaximized
         .Connect = glogon.ConectRPT
         .WindowTitle = "Reporte consulta de movimiento de actividades"
         .ReportFileName = SIFGlobal.fxPathReportes("Comites_ControlLiquidacionEspecifico.rpt")
         
         .Formulas(0) = "fxTitulo= 'CONTROL DE LIQUIDACIONES POR COMITE'"
         .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
         .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
         .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
         
         strSQL = "({afi_cd_cuentas.cod_comite}) = '" & txtCodigoComite.Text & "'"
         
         .SelectionFormula = strSQL
         
         .PrintReport
      End With
End Select

Exit Sub

vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnDetallarLiquidacion_Click()
On Error GoTo vError

Dim vTotal As Double
Dim strSQL As String
Dim i As Integer

    If Trim(txtTotal.Text) = "" Then txtTotal.Text = 0
    
    vTotal = CDbl(txtTotal.Text)
    
    With vGrid
     For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
         .Col = 2
         vOperacion = .Text
            
         'D= Detalle
         strSQL = "update afi_cd_cuentas set PROCESO = 'D'" _
                & " where NOPERACION = '" & vOperacion & "'"
         Call ConectionExecute(strSQL)
            
        End If
        
     Next i
    End With
    Call sbCuentaOpPendientes
    Call sbCargaOperaciones
    Call sbCargaOpxDetallar
    
    txtTotal.Text = Format(vTotal, "Standard")
    
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnLiquidar_Click()
Dim i As Integer
Dim vTipoDoc As String, vTransaccion As String
 
If vOperacion = "" Or vOperacion = "0" Then Exit Sub
 
If CCur(txtDiferencia.Text) > 0 Then
   MsgBox "Existen Diferencias en el detalle con el monto a Cancelar...Revise!", vbExclamation
   Exit Sub
End If
 
With vGridOpxDetallar
  For i = 1 To .MaxRows
   .Row = i
   .Col = 1
   If .Value = vbChecked Then
      .Col = 2
      vOperacion = .Text
      strSQL = "exec spAFI_CD_AsientoLiquidacion " & vOperacion & ",'" & glogon.Usuario & "','" _
             & GLOBALES.gOficinaTitular & "','" & txtNotas.Text & "'"
      
      Call OpenRecordSet(rs, strSQL)
        vTipoDoc = rs!TipoDoc
        vTransaccion = rs!transaccion
      rs.Close
   End If
Next i

Call sbImprimeRecibo(vTransaccion, vTipoDoc)

MsgBox "La Liquidación fue registrada satisfactoriamente", vbInformation, "Información"

Call sbCargaOpxDetallar

End With

End Sub

Private Sub Form_Activate()
 vModulo = 40
End Sub

Private Sub Form_Load()
  vModulo = 40
  
  txtTotal.Text = 0
  txtTotal.Text = 0
  txtTotalFactura.Text = 0
  txtDiferencia.Text = 0
 
  tcMain(0).Selected = True
  
  vGrid.MaxRows = 0
  vGridOpxDetallar.MaxRows = 0
  vGridFacturas.MaxRows = 0
  vGridFacturas.MaxCols = 5
  vGridHistorico.MaxRows = 0
  If GLOBALES.gTag <> Empty Then txtCodigoComite.Text = GLOBALES.gTag
  GLOBALES.gTag = ""
End Sub

Private Sub sbGuardaFactura()
  
  strSQL = "INSERT AFI_CD_DETALLE_LIQUIDACION (NOPERACION, NDOCUMENTO,DEPOSITO, DETALLE, " _
         & "FECHA_DOCUMENTO, MONTO,REGISTRO_FECHA, REGISTRO_USUARIO) Values " _
         & "(" & vOperacion & ",'" & vDocumento & "','" & vDeposito & "','" & vDetalle & "' " _
         & ", '" & Format(vFecha, "yyyymmdd") & "'," & vMonto & ",getdate(),'" & glogon.Usuario & "')"
  Call ConectionExecute(strSQL)
   
End Sub

Private Sub sbModificaFactura()

  strSQL = "UPDATE AFI_CD_DETALLE_LIQUIDACION SET DETALLE ='" & vDetalle & "',FECHA_DOCUMENTO ='" & Format(vFecha, "yyyymmdd") & "'" _
         & ", MONTO=" & vMonto & ",REGISTRO_FECHA =getdate(),REGISTRO_USUARIO ='" & glogon.Usuario & "'" _
         & " WHERE NOPERACION=" & vOperacion & " and NDOCUMENTO ='" & vDocumento & "' "
  Call ConectionExecute(strSQL)

End Sub

Private Sub sbEliminaFactura()
  
  strSQL = "DELETE FROM AFI_CD_DETALLE_LIQUIDACION WHERE NOPERACION=" & vOperacion & " and NDOCUMENTO ='" & vDocumento & "' "
  Call ConectionExecute(strSQL)
  
End Sub
'Trae los Montos de la liquidación
'y el monto total en facturas registradas
Private Sub sbTraeMontos(ByVal vOperacion As Integer)

Dim Saldo, MontoTotal, MontoFacturas As Double
Dim TesoreriaSolucitud As Long
On Error GoTo vError

txtTotal.Text = 0
txtTotalFactura.Text = 0
txtDiferencia.Text = 0

'Se obtiene el Monto Total de la Operacion
strSQL = "Select Monto,NOPERACION from AFI_CD_CUENTAS " _
       & " Where Noperacion = " & vOperacion & ""
Call OpenRecordSet(rs, strSQL)

MontoTotal = Format(rs!Monto, "standard")

rs.Close

MontoFacturas = 0
'Se obtiene el total de las facturas que respaldan la Liquidación
strSQL = "Select isnull(MONTO,0) as 'Monto' from AFI_CD_DETALLE_LIQUIDACION " _
       & " Where NOPERACION = " & vOperacion & " "
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  MontoFacturas = MontoFacturas + Format(rs!Monto, "standard")
  rs.MoveNext
Loop


rs.Close
          
Saldo = MontoTotal - MontoFacturas
  
txtTotal.Text = Format(CDbl(txtTotal.Text) + MontoTotal, "standard")
txtTotalFactura.Text = Format(MontoFacturas, "standard")
txtDiferencia.Text = Format(Saldo, "standard")

Exit Sub
vError:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub sbActualizaSaldo(ByVal vOperacion As Integer)
Dim TotalSaldo As Double
On Error GoTo vError

strSQL = "Select isnull(C.Monto - sum(DL.MONTO),0) as 'Saldo'" _
       & " from dbo.AFI_CD_CUENTAS C" _
       & "  inner join AFI_CD_DETALLE_LIQUIDACION DL on C.NOPERACION = DL.NOPERACION" _
       & " Where C.Noperacion = " & vOperacion
          
Call OpenRecordSet(rs, strSQL)

TotalSaldo = Format(rs!Saldo, "Standard")

rs.Close

strSQL = "Update AFI_CD_CUENTAS set SALDO = " & CCur(TotalSaldo) & ",ESTADO ='L'" _
       & " where NOPERACION= " & vOperacion & ""
Call ConectionExecute(strSQL)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    btnDetallarLiquidacion.Visible = False
    
    Select Case Item.Index
      Case 0
            btnDetallarLiquidacion.Visible = True
      Case 1
        Call sbCargaOpxDetallar
      Case 2
        Call sbCargaHistorico
     End Select

End Sub

Private Sub txtCodigoComite_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
      Call sbCargaComites
      Call sbCuentaOpPendientes
      Call sbCargaOperaciones
    ElseIf KeyCode = vbKeyF4 Then
        gBusquedas.Columna = "C.COD_COMITE"
        gBusquedas.Orden = "C.COD_COMITE"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select C.COD_COMITE,CM.DESCRIPCION" _
                            & " from AFI_CD_CUENTAS C" _
                            & " inner join AFI_CD_COMITES CM on c.COD_COMITE = CM.COD_COMITE and C.ESTADO ='T' "
        frmBusquedas.Show vbModal
        txtCodigoComite = gBusquedas.Resultado
        txtDescripcionComite = gBusquedas.Resultado2
        
        Call sbCuentaOpPendientes
        Call sbCargaOperaciones
    End If
     tcMain(0).Selected = True
     
     GLOBALES.gTag = Trim(txtCodigoComite.Text)

End Sub

Private Sub sbCuentaOpPendientes()
 Dim strSQL As String, rs As New ADODB.Recordset
    strSQL = "select count(COD_COMITE)as Cuenta " _
           & "from AFI_CD_CUENTAS where estado='T' and COD_COMITE='" & Trim(txtCodigoComite.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    txtRate.Text = "Liquidaciones Pendientes : " & IIf(IsNull(rs!Cuenta), 0, rs!Cuenta)
     
    rs.Close
 
End Sub

Private Sub sbCargaComites()
 Dim strSQL As String, rs As New ADODB.Recordset
 
  strSQL = "select DESCRIPCION from AFI_CD_COMITES where COD_COMITE='" & Trim(txtCodigoComite.Text) & "'"
  Call OpenRecordSet(rs, strSQL)
      
  If Not rs.EOF Then
    txtDescripcionComite.Text = Trim(rs!Descripcion)
  End If
  
  rs.Close
End Sub

Private Sub sbCargaOpxDetallar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

vGridFacturas.MaxRows = 0
vGridFacturas.MaxRows = 1


  With vGridOpxDetallar
   .MaxRows = 1
   
   strSQL = "select C.NOPERACION as Operacion,datediff(d,C.REGISTRO_FECHA,getdate()) as 'Dias_Pendientes' " _
         & ", sum(CA.MONTO) as Monto from dbo.AFI_CD_CUENTAS C" _
         & " inner join AFI_CD_CUENTAS_ACTIVIDADES CA on C.NOPERACION = CA.NOPERACION" _
         & " where  C.COD_COMITE = '" & txtCodigoComite.Text & "' and C.Estado = 'T' and C.PROCESO ='D'" _
         & " group by C.NOPERACION,C.REGISTRO_FECHA"
         

  Call OpenRecordSet(rs, strSQL)
  
  Do While Not rs.EOF
    .Row = .MaxRows
    
    .Col = 2
    .Text = rs!Operacion
    
    .Col = 3
    .Text = Format(rs!Monto, "Standard")
    
    .Col = 4
    .Text = rs!Dias_Pendientes
                
    .MaxRows = .MaxRows + 1
    rs.MoveNext
    
  Loop
  .MaxRows = .MaxRows - 1
  rs.Close
  
 End With

Exit Sub

vError:
      MsgBox Err.Description, vbCritical

End Sub

Private Sub sbCargaOperaciones()
On Error GoTo error
Dim strSQL As String, rs As New ADODB.Recordset
Dim vItem As MSComctlLib.ListItem
Dim vLvw As MSComctlLib.ListView
Dim vKey As String
Dim i As Integer

  
   strSQL = "select C.NOPERACION,C.ACTIVA_FECHA , DATEDIFF(D,C.ACTIVA_FECHA,GETDATE()) as 'Dias_Pendientes' " _
         & ",CA.MONTO as Monto,A.DESCRIPCION ACTIVIDAD,case C.ESTADO when 'T'  then 'Trasladado' when 'A'  then 'Activo' " _
         & "Else 'Liquidado' End as Estado,case C.TIPO when 'T' then 'Transferencia' else 'Cheque' End as Desembolso " _
         & ",C.REGISTRO_USUARIO, Tes.FECHA_EMISION" _
         & " from dbo.AFI_CD_CUENTAS C" _
         & " inner join AFI_CD_CUENTAS_ACTIVIDADES CA on C.NOPERACION = CA.NOPERACION" _
         & " inner join AFI_CD_ACTIVIDADES A on CA.COD_ACTIVIDAD = A.COD_ACTIVIDAD" _
         & "  left join TES_TRANSACCIONES Tes on C.TESORERIA_NSOLICITUD = Tes.NSOLICITUD" _
         & " where  C.COD_COMITE='" & txtCodigoComite.Text & "' and C.Estado='T' and C.PROCESO='T'"
  Call OpenRecordSet(rs, strSQL)
  
  With vGrid
    .MaxRows = 1
    
    For i = 1 To .MaxCols
     .Col = i
     .Text = ""
    Next i
          
    Do While Not rs.EOF
      .Row = .MaxRows
      
      .Col = 2
      .Text = IIf(IsNull(rs!Noperacion), 0, rs!Noperacion)
                
      .Col = 3
      .Text = rs!ACTIVA_FECHA & ""
      
      .Col = 4
      .Text = rs!FECHA_EMISION & ""
      
      .Col = 5
      .Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "standard")
      
      .Col = 6
      .Text = rs!ACTIVIDAD
      
      .Col = 7
      .Text = rs!Dias_Pendientes
      
      .Col = 8
      .Text = rs!Estado
      
      .Col = 9
      .Text = rs!desembolso
      
      .Col = 10
      .Text = rs!Registro_Usuario
      
      rs.MoveNext
     .MaxRows = .MaxRows + 1
    
    Loop
    
    rs.Close
   .MaxRows = .MaxRows - 1
     
  End With

Exit Sub

error:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Function fxGridSumaFacturas(vGrid As Object, Columna As Long) As Double
' Este procedimiento valida que solo pueda haber una registro marcado en el grid
Dim suma As Double, i As Long
    

On Error GoTo vError

    suma = 0
    vGrid.Row = 1
    vGrid.Col = 1
    For i = 1 To vGrid.MaxRows
      vGrid.Row = i
      vGrid.Col = Columna
      If IsNumeric(vGrid.Value) Then
         suma = suma + vGrid.Value
      End If
         vGrid.Col = 1
    Next i
    fxGridSumaFacturas = suma
    
Exit Function

vError:
        MsgBox Err.Description
    
End Function

'Modifica las facturas y trae el monto
Private Sub vGridFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

With vGridFacturas
  .Row = .ActiveRow
  
  If .ActiveCol = .MaxCols And KeyCode = vbKeyReturn Then
    .Col = 1
    If Not fxValidaFactura Then
       Call sbGuardaFactura
       .MaxRows = .MaxRows + 1
    Else
       Call sbModificaFactura
       Call sbTraeFacturas
    End If

  End If
  
 If .ActiveCol = .MaxCols And KeyCode = vbKeyDelete Then
     If fxValidaFactura Then
        Call sbEliminaFactura
        Call sbTraeFacturas
     End If
 End If

End With
  
Call sbTraeMontos(vOperacion)
  
Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Function fxValidaFactura() As Boolean
On Error GoTo error

 If vOperacion = Empty Then
    MsgBox "Debe seleccionar una operación"
    Exit Function
 End If
 
 
 With vGridFacturas
    
    .Row = .ActiveRow
    
    .Col = 1
    vDeposito = .Value
    
    .Col = 2
    If .Text = Empty Then
      MsgBox "Falta número de documento"
    Else
      vDocumento = .Text
    End If
    
    .Col = 3
    If .Text = Empty Then
      MsgBox "Falta Fecha de documento"
    Else
      vFecha = .Text
    End If
    
    .Col = 4
    If .Text = Empty Then
      MsgBox "Falta detalle de documento"
    Else
      vDetalle = .Text
    End If
    
    .Col = 5
    If .Text = Empty Then
      MsgBox "Falta monto de documento"
    Else
      vMonto = CCur(.Text)
      .Text = Format(.Text, "standard")
    End If
    
    strSQL = "SELECT NOPERACION,NDOCUMENTO" _
           & " FROM AFI_CD_DETALLE_LIQUIDACION" _
           & " where NOPERACION = " & vOperacion & " and NDOCUMENTO='" & Trim(vDocumento) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs.EOF = True Then
      fxValidaFactura = False
    ElseIf rs!Noperacion = vOperacion And vDocumento = Trim(vDocumento) Then
       fxValidaFactura = True
    Else
       
    End If

    rs.Close
    
 End With


Exit Function
error:
     MsgBox Err.Description, vbCritical
     
End Function

Private Sub sbCargaHistorico()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError



  strSQL = "select C.NOPERACION as Operacion,C.NOTAS,C.LIQUIDA_FECHA,C.ACTIVA_FECHA,Tes.FECHA_EMISION" _
         & ", CA.MONTO as Monto,A.DESCRIPCION ACTIVIDAD,C.TESORERIA_FECHA,C.TESORERIA_NSOLICITUD" _
         & ", case C.ESTADO when 'T'  then 'Trasladado' when 'A'  then 'Activo' " _
         & " Else 'Liquidado' End as Estado ,Case C.APRUEBA when 'J' then 'junta Directiva' when 'O' then 'Oficina Regional' Else" _
         & "' Director Zona' End as Aprueba,case C.TIPO when 'T' then 'Transferencia' else 'Cheque' End as Desembolso,C.REGISTRO_FECHA,C.REGISTRO_USUARIO" _
         & ", Tes.Beneficiario as 'TESORERIA_BENEFICIARIO', Tes.Codigo AS 'TESORERIA_CODIGO'" _
         & " from dbo.AFI_CD_CUENTAS C " _
         & " inner join AFI_CD_CUENTAS_ACTIVIDADES CA on C.NOPERACION = CA.NOPERACION" _
         & " inner join AFI_CD_ACTIVIDADES A on CA.COD_ACTIVIDAD = A.COD_ACTIVIDAD" _
         & " left join TES_Transacciones Tes on C.TESORERIA_NSOLICITUD = Tes.Nsolicitud" _
         & " where  C.COD_COMITE='" & txtCodigoComite.Text & "' order by C.REGISTRO_FECHA desc"
  Call OpenRecordSet(rs, strSQL)
  
  With vGridHistorico
    .MaxRows = 1
    .Row = .MaxRows
         
    Do While Not rs.EOF
      .Row = .MaxRows
      
      .Col = 3
      .Text = IIf(IsNull(rs!Operacion), 0, rs!Operacion)
      
      .Col = 4
      .Text = rs!ACTIVIDAD
      
      .Col = 5
      .Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "standard")
      
      .Col = 6
      .Text = rs!Estado
      
      .Col = 7
      .Text = rs!TESORERIA_FECHA & ""
      
      .Col = 8
      .Text = rs!Registro_Usuario & ""
      
      .Col = 9 'Fecha Activacion
      .Text = rs!ACTIVA_FECHA & ""
      
      .Col = 10 'Fecha Liquidacion
      .Text = rs!LIQUIDA_FECHA & ""
      
      .Col = 11 'No. Solicitud Tesoreria
      .Text = rs!TESORERIA_NSOLICITUD & ""
      
      .Col = 12 'Fecha de Pago Real en Tesorería
      .Text = rs!FECHA_EMISION & ""
      
      .Col = 13 'Beneficiario Tesoreria
      .Text = rs!TESORERIA_CODIGO & ""
      
      .Col = 14 'Beneficiario Tesoreria
      .Text = rs!TESORERIA_BENEFICIARIO & ""
      
      .Col = 15 'notas
      .Text = rs!NOTAS & ""
      
      rs.MoveNext
     .MaxRows = .MaxRows + 1
    Loop
   .MaxRows = .MaxRows - 1
     
  End With

  rs.Close
  
Exit Sub

vError:
      MsgBox Err.Description, vbCritical

End Sub



Private Sub vGridHistorico_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String
Dim vOpRef As String

With vGridHistorico
  .Row = Row
  .Col = 3
  vOpRef = .Text
  vNumDoc = vOpRef
  
  If Col = 1 Then
     vTipoDoc = "CD.CxC"
  Else
     vTipoDoc = "CD.Liq"
     strSQL = "select cod_Transaccion from sif_transacciones" _
            & " where Tipo_Documento = '" & vTipoDoc & "' and Referencia_01 = '" & txtCodigoComite.Text _
            & "' and Referencia_02 = '" & vOpRef & "'"
     Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF And Not rs.BOF Then
         vNumDoc = rs!Cod_Transaccion
     End If
     rs.Close
  End If
  
End With

Call sbImprimeRecibo(vNumDoc, vTipoDoc)

End Sub

Private Sub vGridOpxDetallar_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo error
Dim i As Integer
Dim vTotal As Double

txtTotal.Text = 0
txtDiferencia.Text = 0
vTotal = 0

With vGridOpxDetallar
   For i = 1 To .MaxRows
     .Row = i
     .Col = 1
     If .Value = 1 Then
        .Col = 2
        vOperacion = .Text

        Call sbTraeFacturas
        Call sbTraeMontos(vOperacion)
     End If
   Next i
   

   
End With

Exit Sub
error:
  MsgBox Err.Description
  
End Sub

Private Sub sbTraeFacturas()
Dim vTotalFact As Double
Dim i As Integer
txtTotalFactura = 0
vTotalFact = 0

With vGridFacturas
 .MaxRows = 0
 strSQL = "Select DEPOSITO,NDOCUMENTO,FECHA_DOCUMENTO,DETALLE, MONTO " _
        & "from AFI_CD_DETALLE_LIQUIDACION where NOPERACION = " & vOperacion & " "
 Call OpenRecordSet(rs, strSQL)
 
 .MaxRows = 1
 
 Do While Not rs.EOF
    
    .Row = .MaxRows
    .Col = 1
    .Value = rs!DEPOSITO
    
    .Col = 2
    .Text = rs!nDocumento
    
    .Col = 3
    .Text = Format(rs!FECHA_DOCUMENTO, "dd/mm/yyyy")
    
    .Col = 4
    .Text = rs!Detalle
    
    .Col = 5
    .Text = Format(rs!Monto, "standard")
    vTotalFact = CDbl(vTotalFact) + Format(rs!Monto, "standard")
    
    rs.MoveNext
    .MaxRows = .MaxRows + 1
 
 Loop
 rs.Close
End With

End Sub


