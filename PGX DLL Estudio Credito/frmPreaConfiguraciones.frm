VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPreaConfiguraciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estudio de Crédito: Configuraciones"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   12495
      _Version        =   1572864
      _ExtentX        =   22040
      _ExtentY        =   12938
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
      ItemCount       =   5
      Item(0).Caption =   "Comités"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tcComites"
      Item(1).Caption =   "Garantías"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tcGarantias"
      Item(2).Caption =   "Cambio de Estado"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gEstado"
      Item(3).Caption =   "Monitoreo Proceso Automático"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "GroupBox1"
      Item(4).Caption =   "Edad de Pensión"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "gPension"
      Item(4).Control(1)=   "txtEP_Codigo"
      Item(4).Control(2)=   "txtEP_Descripcion"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2655
         Left            =   -66880
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   6135
         _Version        =   1572864
         _ExtentX        =   10821
         _ExtentY        =   4683
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton btnMonitoreo 
            Height          =   495
            Left            =   4560
            TabIndex        =   13
            Top             =   1920
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Generar"
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
            Picture         =   "frmPreaConfiguraciones.frx":0000
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInicio 
            Height          =   330
            Left            =   1560
            TabIndex        =   16
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpCorte 
            Height          =   330
            Left            =   1560
            TabIndex        =   17
            Top             =   1200
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Datos del Reporte"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   15
            Top             =   1200
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha Corte"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   720
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha Inicio"
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
      Begin XtremeSuiteControls.TabControl tcGarantias 
         Height          =   6855
         Left            =   -70000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   12495
         _Version        =   1572864
         _ExtentX        =   22040
         _ExtentY        =   12091
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
         Item(0).Caption =   "% Liquidez Mínima"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gGarantia"
         Item(1).Caption =   "Garantía Refunde"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "gGarantiaRefunde"
         Begin FPSpreadADO.fpSpread gGarantia 
            Height          =   6015
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   9975
            _Version        =   524288
            _ExtentX        =   17595
            _ExtentY        =   10610
            _StockProps     =   64
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
            MaxCols         =   3
            RowHeaderDisplay=   0
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaConfiguraciones.frx":0719
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread gGarantiaRefunde 
            Height          =   6255
            Left            =   -69880
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   12255
            _Version        =   524288
            _ExtentX        =   21616
            _ExtentY        =   11033
            _StockProps     =   64
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
            RowHeaderDisplay=   0
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaConfiguraciones.frx":0D3D
            AppearanceStyle =   1
         End
      End
      Begin XtremeSuiteControls.TabControl tcComites 
         Height          =   6975
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   12495
         _Version        =   1572864
         _ExtentX        =   22040
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
         Item(0).Caption =   "Máximo Permitido Comité"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gComiteMax"
         Item(1).Caption =   "Líneas Permitidas Monto Máximo"
         Item(1).ControlCount=   3
         Item(1).Control(0)=   "gLineas"
         Item(1).Control(1)=   "txtLP_Codigo"
         Item(1).Control(2)=   "txtLP_Descripcion"
         Item(2).Caption =   "Obligatoriedad Adjuntos"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "gAdjuntos"
         Begin FPSpreadADO.fpSpread gComiteMax 
            Height          =   6495
            Left            =   0
            TabIndex        =   3
            Top             =   360
            Width           =   12495
            _Version        =   524288
            _ExtentX        =   22040
            _ExtentY        =   11456
            _StockProps     =   64
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
            MaxCols         =   7
            RowHeaderDisplay=   0
            SpreadDesigner  =   "frmPreaConfiguraciones.frx":581E
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread gLineas 
            Height          =   5775
            Left            =   -68920
            TabIndex        =   4
            Top             =   960
            Visible         =   0   'False
            Width           =   9735
            _Version        =   524288
            _ExtentX        =   17171
            _ExtentY        =   10186
            _StockProps     =   64
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
            MaxCols         =   3
            RowHeaderDisplay=   0
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaConfiguraciones.frx":6061
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtLP_Codigo 
            Height          =   375
            Left            =   -68920
            TabIndex        =   5
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   661
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
         Begin FPSpreadADO.fpSpread gAdjuntos 
            Height          =   6255
            Left            =   -68800
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   9975
            _Version        =   524288
            _ExtentX        =   17595
            _ExtentY        =   11033
            _StockProps     =   64
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
            MaxCols         =   3
            RowHeaderDisplay=   0
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaConfiguraciones.frx":66A3
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtLP_Descripcion 
            Height          =   375
            Left            =   -67480
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   8295
            _Version        =   1572864
            _ExtentX        =   14631
            _ExtentY        =   661
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
      End
      Begin FPSpreadADO.fpSpread gPension 
         Height          =   6495
         Left            =   -70000
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   12495
         _Version        =   524288
         _ExtentX        =   22040
         _ExtentY        =   11456
         _StockProps     =   64
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
         MaxCols         =   6
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaConfiguraciones.frx":6CD8
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gEstado 
         Height          =   6495
         Left            =   -69880
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   12255
         _Version        =   524288
         _ExtentX        =   21616
         _ExtentY        =   11456
         _StockProps     =   64
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
         MaxCols         =   7
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaConfiguraciones.frx":7495
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtEP_Codigo 
         Height          =   375
         Left            =   -68680
         TabIndex        =   20
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin XtremeSuiteControls.FlatEdit txtEP_Descripcion 
         Height          =   375
         Left            =   -67240
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   8295
         _Version        =   1572864
         _ExtentX        =   14631
         _ExtentY        =   661
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimientos de Estudio de Créditos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaConfiguraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, mTipo As String

Private Sub sbLista(pTipo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

mTipo = UCase(pTipo)

vPaso = True

Select Case pTipo
    Case "ComiteMax"
            strSQL = "exec spCrdPreaListaParametrosComite"
            Call sbCargaGrid(gComiteMax, gComiteMax.MaxCols, strSQL)
            gComiteMax.MaxRows = gComiteMax.MaxRows - 1
    
    Case "ComiteLineas"
            If txtLP_Codigo.Text = "" Then
                strSQL = "exec spCrdPreaListaConsultaParametrosValidaMaximoP '', 'N'"
            Else
                strSQL = "exec spCrdPreaListaConsultaParametrosValidaMaximoP '" & txtLP_Codigo.Text & "', 'S'"
            End If
            
            Call sbCargaGrid(gLineas, gLineas.MaxCols, strSQL)
            gLineas.MaxRows = gLineas.MaxRows - 1
    
    Case "ComiteAdjuntos"
    
            strSQL = "exec spCrdPreaListaComitesParametrizacionAdj"
            Call sbCargaGrid(gAdjuntos, gAdjuntos.MaxCols, strSQL)
            gAdjuntos.MaxRows = gAdjuntos.MaxRows - 1
    
    '---Garantias
    Case "GarantiaLiqMin"
            strSQL = "exec spCrdPreaListaParametrosGarantiaLiquido"
            Call sbCargaGrid(gGarantia, gGarantia.MaxCols, strSQL)
            gGarantia.MaxRows = gGarantia.MaxRows - 1
    
    Case "GarantiaRefunde"
            strSQL = "exec spCrdPreaListaGarantiasRefunde"
            Call sbCargaGrid(gGarantiaRefunde, gGarantiaRefunde.MaxCols, strSQL)
            gGarantiaRefunde.MaxRows = gGarantiaRefunde.MaxRows - 1

    '--Cambios de Estado
    
    Case "Estado"
            strSQL = "exec spCrdPreaListaMotivosCambioEstado"
            Call sbCargaGrid(gEstado, gEstado.MaxCols, strSQL)


    
    '--Edad de Pension
    Case "Pension"
            If txtLP_Codigo.Text = "" Then
                strSQL = "exec spCrdPreaListaLineasConfigEdadPension '', 'N'"
            Else
                strSQL = "exec spCrdPreaListaLineasConfigEdadPension '" & txtEP_Codigo.Text & "', 'S'"
            End If
            Call sbCargaGrid(gPension, gPension.MaxCols, strSQL)
            gPension.MaxRows = gPension.MaxRows - 1
    


End Select

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

'Consulta Inicial
tcMain.Item(0).Selected = True
tcComites.Item(0).Selected = True
Call sbLista("ComiteMax")

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub gAdjuntos_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pId As String, pDetalle As String, pValor As Byte

With gAdjuntos
  .Row = Row
  .Col = 1
  pId = .Text
  .Col = 2
  pDetalle = .Text
  .Col = 3
  pValor = .Value
  .Col = Col
  pDetalle = "Config: Comité [" & pDetalle & "] Obligatorio Adjuntos [" & IIf((.Value), "Sí", "No") & "]"
End With

strSQL = "exec spCrdPreaGuardaParametroComiteAdjuntoOblig " & pId & ", " & pValor
Call ConectionExecute(strSQL)

Call Bitacora("Registra", pDetalle)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub gComiteMax_KeyDown(KeyCode As Integer, Shift As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

If KeyCode = vbKeyReturn And gComiteMax.ActiveCol = gComiteMax.MaxCols Then

    Me.MousePointer = vbHourglass
    
    Dim pId As String, pDetalle As String, pAhorros As Currency, pPagare As Currency
    Dim pHipotecario As Currency, pPrendario As Currency, pFiduciaria As Currency

    
    With gComiteMax
      .Row = .ActiveRow
      .Col = 1
      pId = .Text
      .Col = 2
      pDetalle = .Text
      .Col = 3
      pAhorros = CCur(.Text)
      .Col = 4
      pPagare = CCur(.Text)
      .Col = 5
      pHipotecario = CCur(.Text)
      .Col = 6
      pPrendario = CCur(.Text)
      .Col = 7
      pFiduciaria = CCur(.Text)
    
    End With
    
    
    pDetalle = "Config: Comité [" & pId & "] Máximos Garantia > [Ah: " & Format(pAhorros, "Standard") & "] [Pag: " _
             & Format(pPagare, "Standard") & "] [Hip: " & Format(pHipotecario, "Standard") & "] [Pren: " _
             & Format(pPrendario, "Standard") & "] [Fidu: " & Format(pFiduciaria, "Standard") & "]"
    
    strSQL = "exec spCrdPreaGuardaParametroComite " & pId & ", " & pAhorros & ", " & pPagare & ", " & pHipotecario _
            & ", " & pPrendario & ", " & pFiduciaria
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", pDetalle)
    
    Me.MousePointer = vbDefault
    
    MsgBox "Montos Máximos por Garantías, actualizados satisfactoriamente!", vbInformation
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub gEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

If KeyCode = vbKeyReturn And gEstado.ActiveCol = 3 Then

    Me.MousePointer = vbHourglass
    
    Dim pId As Long, pDetalle As String, pActivo As Integer
    
    With gEstado
      .Row = .ActiveRow
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pDetalle = .Text
      .Col = 3
      If .Value = True Or .Value = False Then
         pActivo = IIf(.Value, 1, 0)
      Else
        pActivo = .Value
      End If

    End With
    
    strSQL = "exec spCrdPreaGuardaParametroMotivoCambioEstado " & pId & ", '" & pDetalle & "', " & pActivo _
           & ", '" & glogon.Usuario & "', '" & IIf(pId = 0, "R", "M") & "'"
    Call ConectionExecute(strSQL)
    
    pDetalle = "Config: Cambio Estado [Id: " & pId & "..." & pDetalle & "] Activo: [" & IIf(pActivo = 1, "Sí", "No") & "]"
    Call Bitacora(IIf(pId = 0, "Registra", "Modifica"), pDetalle)
    
    Me.MousePointer = vbDefault

    MsgBox pDetalle & ">>>> " & IIf(pId = 0, "Registrada", "Modificada") & " satisfactoriamente!", vbInformation
    Call sbLista("Estado")
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub gGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

If KeyCode = vbKeyReturn And gGarantia.ActiveCol = gGarantia.MaxCols Then

    Me.MousePointer = vbHourglass
    
    Dim pId As String, pDetalle As String, pValor As Currency
    
    With gGarantia
      .Row = .ActiveRow
      .Col = 1
      pId = .Text
      .Col = 2
      pDetalle = .Text
      .Col = 3
      pValor = .Value
    End With
    
    pDetalle = "Config: Garantía [" & pDetalle & "] % Liquidez Mínima: " & Format(pValor, "Standard") & "%"
    
    strSQL = "exec spCrdPreaGuardaParametroGarantiaLiquido '" & pId & "', " & pValor
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", pDetalle)
    
    Me.MousePointer = vbDefault

    MsgBox "Porcentaje de Liquidez Mínima por Garantía, actualizada satisfactoriamente!", vbInformation

End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub gGarantiaRefunde_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pIdGarantia As String, pDetalle As String
Dim pAhorro As Byte, pPrendario As Byte, pHipotecario As Byte
Dim pFiduciario As Byte, pPagare As Byte, pExcedente As Byte

With gGarantiaRefunde
  .Row = Row
  .Col = 1
  pIdGarantia = .Text
  .Col = 2
  pDetalle = .Text
  .Col = 3
  pAhorro = .Value
  .Col = 4
  pPrendario = .Value
  .Col = 5
  pHipotecario = .Value
  .Col = 6
  pFiduciario = .Value
  .Col = 7
  pPagare = .Value
  .Col = 8
  pExcedente = .Value


  .Col = Col
  Select Case Col
    Case 3
        pDetalle = "Config: Garantía [" & pDetalle & "] Refunde Garantía [" & IIf((.Value), "Sí", "No") & "] Sobre Ahorros"
    Case 4
        pDetalle = "Config: Garantía [" & pDetalle & "] Refunde Garantía [" & IIf((.Value), "Sí", "No") & "] Prendaria"
    Case 5
        pDetalle = "Config: Garantía [" & pDetalle & "] Refunde Garantía [" & IIf((.Value), "Sí", "No") & "] Hipotecaria"
    Case 6
        pDetalle = "Config: Garantía [" & pDetalle & "] Refunde Garantía [" & IIf((.Value), "Sí", "No") & "] Fiduciaria"
    Case 7
        pDetalle = "Config: Garantía [" & pDetalle & "] Refunde Garantía [" & IIf((.Value), "Sí", "No") & "] Pagaré"
    Case 8
        pDetalle = "Config: Garantía [" & pDetalle & "] Refunde Garantía [" & IIf((.Value), "Sí", "No") & "] S/Excedentes"
  End Select
End With

strSQL = "exec spCrdPreaGuardaGarantiaRefunde '" & pIdGarantia & "', " & pAhorro & ", " & pPrendario & ", " & pHipotecario _
       & ", " & pFiduciario & ", " & pPagare & ", " & pExcedente
Call ConectionExecute(strSQL)

Call Bitacora("Registra", pDetalle)




Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub gLineas_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pId As String, pDetalle As String, pValor As Byte

With gLineas
  .Row = Row
  .Col = 1
  pId = .Text
  .Col = 2
  pDetalle = .Text
  .Col = 3
  pValor = .Value
  
  .Col = Col
  pDetalle = "Config: Líneas [" & pDetalle & "] Valida Monto Máximo [" & IIf((.Value), "Sí", "No") & "]"
End With

strSQL = "exec spCrdPreaGuardaParametroComiteValidaMaximo '" & pId & "', " & pValor
Call ConectionExecute(strSQL)

Call Bitacora("Registra", pDetalle)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub gPension_KeyDown(KeyCode As Integer, Shift As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

If KeyCode = vbKeyReturn And gPension.ActiveCol = gPension.MaxCols Then

    Me.MousePointer = vbHourglass
    
    Dim pId As String, pDetalle As String, pIEstudio As Integer, pIFormaliza As Integer
    Dim pGarantias As String, pComites As String
    
    With gPension
      .Row = .ActiveRow
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pDetalle = .Text
      .Col = 3
      If .Value = True Or .Value = False Then
         pIEstudio = IIf(.Value, 1, 0)
      Else
         pIEstudio = .Value
      End If

      .Col = 4
      If .Value = True Or .Value = False Then
         pIFormaliza = IIf(.Value, 1, 0)
      Else
         pIFormaliza = .Value
      End If
    
      .Col = 5
      pGarantias = Replace(.Text, "Ninguna", "")
      
      .Col = 6
      pComites = Replace(.Text, "Ninguno", "")
    
    End With
                             
    strSQL = "exec spCrdPreaGuardaParametroEdadPension '" & pId & "', '" & pGarantias & "', '" & pComites _
           & "', " & pIEstudio & ", " & pIFormaliza & ", '" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    
    pDetalle = "Config: Edad Pensión [Id: " & pId & "] Aplica: [Estudio: " & IIf(pIEstudio = 1, "Sí", "No") _
             & ", Formaliza: " & IIf(pIFormaliza = 1, "Sí", "No") & "] Garantías: " & pGarantias & ", Comités: " & pComites
    Call Bitacora("Modifica", pDetalle)
    
    Me.MousePointer = vbDefault

    MsgBox pDetalle & ">>>> " & ", actualizada satisfactoriamente!", vbInformation

End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcComites_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Maximo
        Call sbLista("ComiteMax")
    Case 1 'Lineas
        Call sbLista("ComiteLineas")
    Case 2 'Adjuntos
        Call sbLista("ComiteAdjuntos")
End Select

End Sub

Private Sub tcGarantias_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 0 'Liquidez Minima
      Call sbLista("GarantiaLiqMin")
    Case 1 'Garantias Refunde
      Call sbLista("GarantiaRefunde")
End Select
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Comites
      tcComites.Item(0).Selected = True
      Call sbLista("ComiteMax")
      
    Case 1 'Garantias
      tcGarantias.Item(0).Selected = True
      Call sbLista("GarantiaLiqMin")
      
    Case 2 'Estados
      Call sbLista("Estado")
    
    Case 3 'Monitoreo
    
    Case 4 'Edad Pension
      Call sbLista("Pension")


End Select


End Sub


Private Sub txtEP_Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Filtro = " and Poliza = 'N' and Retencion = 'N'"
  frmBusquedas.Show vbModal
  txtEP_Codigo.Text = gBusquedas.Resultado
  txtEP_Descripcion.Text = gBusquedas.Resultado2
  Call sbLista("Pension")
End If

End Sub



Private Sub txtLP_Codigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Filtro = " and Poliza = 'N' and Retencion = 'N'"
  frmBusquedas.Show vbModal
  txtLP_Codigo.Text = gBusquedas.Resultado
  txtLP_Descripcion.Text = gBusquedas.Resultado2
  Call sbLista("ComiteLineas")
End If

End Sub
