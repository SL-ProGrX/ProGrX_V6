VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_CD_Actividades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descripción de Actividades"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   14025
   Icon            =   "FrmAF_CD_Actividades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   13815
      _Version        =   1572864
      _ExtentX        =   24368
      _ExtentY        =   12515
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
      Item(0).Caption =   "Actividades"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Definición de Montos"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "lswActividades"
      Item(1).Control(1)=   "vGridMontos"
      Item(1).Control(2)=   "lblActividad"
      Item(1).Control(3)=   "ShortcutCaption1"
      Item(1).Control(4)=   "btnActividadRefresh"
      Item(2).Caption =   "Consulta"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "dtpInicio"
      Item(2).Control(1)=   "dtpCorte"
      Item(2).Control(2)=   "Label1(0)"
      Item(2).Control(3)=   "Label1(1)"
      Item(2).Control(4)=   "cboActividad"
      Item(2).Control(5)=   "btnBuscar"
      Item(2).Control(6)=   "btnExportar"
      Item(2).Control(7)=   "lsw"
      Item(2).Control(8)=   "ProgressBarX"
      Item(2).Control(9)=   "btnReporte"
      Begin XtremeSuiteControls.ListView lswActividades 
         Height          =   5895
         Left            =   -69880
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1572864
         _ExtentX        =   11668
         _ExtentY        =   10398
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5895
         Left            =   -69880
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1572864
         _ExtentX        =   23945
         _ExtentY        =   10398
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Left            =   -60040
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "FrmAF_CD_Actividades.frx":3482
      End
      Begin XtremeSuiteControls.PushButton btnActividadRefresh 
         Height          =   375
         Left            =   -63760
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "FrmAF_CD_Actividades.frx":3B82
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6480
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Para seleccionar la cuenta contable (F4)"
         Top             =   480
         Width           =   13650
         _Version        =   524288
         _ExtentX        =   24077
         _ExtentY        =   11430
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
         MaxCols         =   7
         ScrollBars      =   2
         SpreadDesigner  =   "FrmAF_CD_Actividades.frx":4282
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridMontos 
         Height          =   5865
         Left            =   -63205
         TabIndex        =   3
         Top             =   975
         Visible         =   0   'False
         Width           =   6915
         _Version        =   524288
         _ExtentX        =   12197
         _ExtentY        =   10345
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
         SpreadDesigner  =   "FrmAF_CD_Actividades.frx":4A24
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   -68920
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   330
         Left            =   -67480
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboActividad 
         Height          =   330
         Left            =   -64480
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Left            =   -58360
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "FrmAF_CD_Actividades.frx":50E0
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   135
         Left            =   -69880
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1572864
         _ExtentX        =   23945
         _ExtentY        =   238
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   1
      End
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   375
         Left            =   -58960
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "FrmAF_CD_Actividades.frx":524A
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   210
         Index           =   1
         Left            =   -65800
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Actividad"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   210
         Index           =   0
         Left            =   -69760
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Fechas"
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
      Begin XtremeShortcutBar.ShortcutCaption lblActividad 
         Height          =   375
         Left            =   -63280
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   6975
         _Version        =   1572864
         _ExtentX        =   12303
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -69880
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1572864
         _ExtentX        =   11668
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Actividades disponibles: Seleccione!"
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
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   582
      ButtonWidth     =   2646
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualiza Año"
            Key             =   "Actualiza"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":5951
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":5A6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":5B79
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":5CA2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAF_CD_Actividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vGuarda As Boolean

Dim Nuevo As Boolean
Dim Filas As Integer, Columnas As Integer, Inc As Integer, Can As Integer
Dim vCodigo As Integer
Dim vActivo As Boolean

Private Function fxConsecutivoCodigoMonto()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select coalesce(max(cod_monto),0) + 1 as Ultimo from afi_cd_actividades_rangos"
Call OpenRecordSet(rs, strSQL)
   fxConsecutivoCodigoMonto = rs!ultimo
rs.Close
End Function

Private Function fxConsecutivo()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull( MAX(CAST(cod_actividad AS INT)) , 0)  + 1 as Ultimo from AFI_CD_ACTIVIDADES"
Call OpenRecordSet(rs, strSQL)
   fxConsecutivo = rs!ultimo
rs.Close

End Function

Private Sub sbConsultaActividades()

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select Act.descripcion, S.descripcion as 'Comite', C.monto, C.noperacion, C.registro_usuario, C.registro_fecha, C.estado " _
         & ", Est.NombreEstado" _
         & " from afi_cd_cuentas C inner join afi_cd_cuentas_actividades Ca on C.noperacion = Ca.noperacion" _
         & " inner join afi_cd_actividades Act on Act.cod_actividad = Ca.cod_actividad " _
         & " inner join afi_cd_comites S on C.Cod_comite = S.cod_comite " _
         & " inner join AFI_CD_TIPOS_ESTADOS_CUENTAS Est on C.Estado = Est.CodEstado " _
         & " where C.registro_fecha between '" & Format(dtpInicio.Value, "yyyymmdd 00:00:00") & "' " _
         & " and '" & Format(dtpCorte.Value, "yyyymmdd 23:59:59") & "'"
If cboActividad.Text <> "TODOS" Then
   strSQL = strSQL & " and Ca.Cod_Actividad = '" & cboActividad.ItemData(cboActividad.ListIndex) & "'"
End If
         
 Call OpenRecordSet(rs, strSQL)
 
 Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , IIf(IsNull(rs!Descripcion), "", rs!Descripcion))
               itmX.SubItems(1) = Trim(IIf(IsNull(rs!Comite), "", rs!Comite))
               itmX.SubItems(2) = Format(Trim(IIf(IsNull(rs!Monto), "0", rs!Monto)), "Standard")
               itmX.SubItems(3) = Trim(IIf(IsNull(rs!Noperacion), "", rs!Noperacion))
               itmX.SubItems(4) = Trim(IIf(IsNull(rs!Registro_Usuario), "", rs!Registro_Usuario))
               itmX.SubItems(5) = Format(Trim(IIf(IsNull(rs!Registro_Fecha), "1990-01-01", rs!Registro_Fecha)), "yyyy-mm-dd hh:mm:ss")
               itmX.SubItems(6) = Trim(IIf(IsNull(rs!NombreEstado), "", rs!NombreEstado))
     rs.MoveNext
 Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbControlLiquidacion()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
 
 .ReportFileName = SIFGlobal.fxPathReportes("vista_afi_cd_cuentasactivas")
 .WindowTitle = "Reporte Actividades"
  strSQL = strSQL & "cdate({vista_afi_cd_cuentasactivas.tesoreria_fecha}) " _
  & "in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
  strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
 .SelectionFormula = strSQL
  
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='Control de Liquidaciones'"
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub sblswActividadesividadesCarga()
 
vGridMontos.MaxRows = 0
 
strSQL = "select cod_actividad,descripcion from afi_cd_actividades order by cod_actividad asc"
Call OpenRecordSet(rs, strSQL)
        
With lswActividades.ListItems
  .Clear
  Do While Not rs.EOF
      Set itmX = .Add(, , rs!Cod_Actividad)
          itmX.SubItems(1) = IIf(IsNull(rs!Descripcion), "Sin Nombre", rs!Descripcion)
      rs.MoveNext
  Loop
    rs.Close
End With

End Sub

Private Sub sbCambiaAno()

Dim fecPeriodo As String
Dim fecLiq As String
Dim i As Integer, S As Integer

S = MsgBox("Desea cambiar el año en curso", vbYesNo + vbInformation, "Información")

If S = vbYes Then
 For i = 1 To vGrid.MaxRows
   vGrid.Row = i
   vGrid.Col = 5
   fecPeriodo = Format(vGrid.Text, "dd/mm/")
   vGrid.Text = fecPeriodo & Year(fxFechaServidor)
   vGrid.Col = 6
   fecLiq = Format(vGrid.Text, "dd/mm/")
   vGrid.Text = fecLiq & Year(fxFechaServidor)
   fecPeriodo = ""
   fecLiq = ""
 Next i
End If

End Sub

Private Sub sbCargaCuenta()

Dim i As Integer

For i = 1 To vGrid.MaxRows

    vGrid.Row = i
    vGrid.Col = 3
    If vGrid.Col = 3 And (vGrid.Text = Empty Or vGrid.Text = "---") Then
        frmCntX_ConsultaCuentas.Show vbModal
        
        vGrid.Col = 3
        vGrid.Text = gCuenta
        
'        vGrid.Col = 3
'        vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text, 0)
        
        vGrid.Col = 4
        vGrid.Text = fxgCntCuentaDesc(gCuenta)
     
     Exit Sub
    End If
 Next i

End Sub


Private Sub sbCargaGridLocal()
Dim i As Integer

strSQL = "select Act.COD_ACTIVIDAD,Act.DESCRIPCION,Act.COD_CUENTA,Act.FECHAPERIOCIDAD,Act.FECHALIQ, Act.ACTIVA,Cta.DESCRIPCION as 'CuentaX'" _
        & ", case when Act.TIPO = 'T' then 'Trimestral' else 'Especial' end as 'Tipo'" _
        & " from AFI_CD_ACTIVIDADES Act left join CNTX_CUENTAS Cta on Act.cod_cuenta = Cta.Cod_cuenta and Cta.cod_contabilidad = " & GLOBALES.gEnlace & "" _
        & "order by Act.COD_ACTIVIDAD asc"

Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    .MaxCols = 7
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      For i = 1 To .MaxCols
         .Col = i
         Select Case i
              Case 1
                vGrid.Text = rs!Cod_Actividad
              Case 2
                vGrid.Text = rs!Descripcion
              Case 3
                .Text = fxgCntCuentaFormato(True, rs!cod_cuenta, 0)
                .TextTip = TextTipFixed
                .TextTipDelay = 1000
                .CellNote = rs!CuentaX & ""
              Case 4
                vGrid.Text = rs!Tipo
              Case 5
                vGrid.Text = Format(rs!fechaperiocidad, "dd/mm/yyyy")
              Case 6
                vGrid.Text = Format(rs!fechaliq, "dd/mm/yyyy")
              Case 7
                vGrid.Value = rs!Activa
         
         End Select
         
      Next i
      
        
        rs.MoveNext
    Loop

End With
rs.Close
vGrid.MaxRows = vGrid.MaxRows + 1


End Sub

Private Sub sbReporte()
On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = ""

With frmContenedor.Crt
 
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
 
 .ReportFileName = SIFGlobal.fxPathReportes("Comites_Actividades.rpt")
 .WindowTitle = "Reporte Actividades"
 
' .SelectionFormula = "{afi_cd_nombramientos.id_pricomite} = '" & LswComi.SelectedItem & "' " _
' & "and {afi_cd_nombramientos_h.estado} = '1'"
  
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='ACTIVIDADES DE COMITES Y DELEGADOS'"
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical


End Sub

Private Sub Cmdimprimir_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
 
 .ReportFileName = SIFGlobal.fxPathReportes("Comites_Actividades.rpt")
 .WindowTitle = "Reporte Actividades y sus montos"
 
' .SelectionFormula = "{afi_cd_nombramientos.id_pricomite} = '" & LswComi.SelectedItem & "' " _
' & "and {afi_cd_nombramientos_h.estado} = '1'"
  
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='ACTIVIDADES Y SUS MONTOS'"
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub btnActividadRefresh_Click()
 Call sblswActividadesividadesCarga
End Sub

Private Sub btnBuscar_Click()
Call sbConsultaActividades
End Sub

Private Sub btnExportar_Click()
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

Private Sub btnReporte_Click()
   strSQL = ""
   With frmContenedor.Crt
      .Reset
      .WindowTitle = "Reporte consulta de movimiento de actividades"
      .WindowShowGroupTree = True
      .WindowShowPrintSetupBtn = True
      .WindowShowRefreshBtn = True
      .WindowShowSearchBtn = True
      .WindowState = crptMaximized
      .Connect = glogon.ConectRPT
      
       strSQL = strSQL & "cdate({AFI_CD_vDesembolsos.TesoreriaFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
       strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
      
      If cboActividad.Text <> "TODOS" Then
        strSQL = " and {AFI_CD_vDesembolsos.TesoreriaFecha}"
      End If
      
      .ReportFileName = SIFGlobal.fxPathReportes("Comites_ControlDesembolsosActividades.rpt")
      
      .Formulas(0) = "fxTitulo= 'CONTROL DE DESEMBOLSOS SEGUN ACTIVIDADES'"
      .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
      .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
      .Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
      .Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
      
      .SelectionFormula = strSQL
        
      .PrintReport
   End With
End Sub

Private Sub Form_Activate()
 vModulo = 40
End Sub

Private Sub Form_Load()
 
 vModulo = 40
  
 vActivo = False
 
 tcMain(0).Selected = True
 
 
 With lswActividades.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", lswActividades.Width - 1400
 End With
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Actividad", 3500
    .Add , , "Comité", 3500
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "N.Operación", 1500, vbCenter
    .Add , , "R.Usuario", 2100, vbCenter
    .Add , , "R.Fecha", 2100
    .Add , , "Estado", 2100, vbCenter
End With

 
 Call Formularios(Me)
 Call RefrescaTags(Me)

 Call sbCargaGridLocal
 
 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value

End Sub


Private Sub OptAct_Click()
 Call sbCargaGridLocal
End Sub

Private Sub OptDes_Click()
 Call sbCargaGridLocal
End Sub



Private Sub vGridact_DblClick(ByVal Col As Long, ByVal Row As Long)
 Call sbCargaCuenta
End Sub


Private Sub vGridact_KeyDown(KeyCode As Integer, Shift As Integer)

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim Conse As Integer, Inc As Integer

strSQL = "select coalesce(max(codtipo),0) + 1 as Ultimo from afi_cd_periocidadactividades"
          Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
 Conse = rs!ultimo
End If
rs.Close

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    Inc = vGrid.MaxRows
    vGrid.InsertRows vGrid.ActiveRow + 1, 1
    vGrid.MaxRows = vGrid.MaxRows
    vGrid.SetActiveCell 0, vGrid.MaxRows
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 1
    vGrid.Text = Conse
    Nuevo = True
    Call sbCargaCuenta
       
End If

If KeyCode = vbKeyDelete Then
  If MsgBox("¿Desea eliminar esta linea?", vbYesNo Or vbQuestion, "") = vbYes Then
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 2
     If vGrid.MaxRows > 0 Then
       vGrid.MaxRows = vGrid.MaxRows - 1
       vGrid.DeleteRows vGrid.ActiveRow + 1, 1
       vGrid.Row = vGrid.ActiveRow
     End If
End If




End If

End Sub

Private Sub vGridact_KeyPress(KeyAscii As Integer)
 Call sbCargaCuenta
End Sub






Private Sub lswActividades_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If lswActividades.ListItems.Count <= 0 Then Exit Sub

vCodigo = Item.Text
lblActividad.Caption = Trim(Item.SubItems(1))

strSQL = " select cod_monto,monto,minimo,maximo from afi_cd_actividades_rangos where cod_actividad = " & vCodigo & ""
Call sbCargaGrid(vGridMontos, 4, strSQL)

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 
Select Case Item.Index
 Case 1
     Call sblswActividadesividadesCarga
 Case 2
    strSQL = "select cod_actividad as 'IdX',descripcion as 'ItmX' from afi_cd_actividades order by cod_actividad asc"
    Call sbCbo_Llena_New(cboActividad, strSQL, True, True)
    
End Select
 
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
  Case "REPORTE"
     Call sbReporte
  Case "ACTUALIZA"
     Call sbCambiaAno
End Select
End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)


Dim i As Long, strSQL As String
Dim rs As New ADODB.Recordset
On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarActividad
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        If vGrid.MaxRows <= vGrid.ActiveRow Then
           vGrid.MaxRows = vGrid.MaxRows + 1
           vGrid.Row = vGrid.MaxRows
        End If
  End If 'Actualiza o Inserta
End If

'Inserta Cuenta
If KeyCode = vbKeyF4 Then
  Call sbCargaCuenta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 1

 If vGrid.Text = "" Then
     vGrid.DeleteRows vGrid.ActiveRow, 1
     vGrid.MaxRows = vGrid.MaxRows - 1
     If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     Exit Sub
 End If
     
   strSQL = "select A.cod_actividad,C.cod_comite,A.descripcion " _
            & "from afi_cd_actividades A left join afi_cd_comites_actividades C " _
            & "on A.cod_actividad = C.cod_actividad where A.cod_actividad = " & vGrid.Text & " and C.cod_comite is not null"
            rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     
          If Not rs.EOF Then
                   MsgBox "Actualmente esta actividad pertenece al comité " & rs!cod_comite & " " & rs!Descripcion & " no podra eliminarla", vbInformation, "Información"
                   rs.Close
                   Exit Sub
          Else
                  
                  i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
                  If i = vbYes Then
                    
                      strSQL = "delete afi_cd_comites_actividades where cod_actividad = " & vGrid.Text
                      Call ConectionExecute(strSQL)
                     
                      strSQL = vGrid.Text
                      vGrid.Col = 2
                      'Call Bitacora("Elimina", "Director: " & vGrid.Text & ")
                     
                      vGrid.DeleteRows vGrid.ActiveRow, 1
                      vGrid.MaxRows = vGrid.MaxRows - 1
                      If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
                     
          End If
          rs.Close
   End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub




Private Sub vGridMontos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String
Dim rs As New ADODB.Recordset
On Error GoTo vError

If vGridMontos.ActiveCol = vGridMontos.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarMontos
  If i > 0 Then
        vGridMontos.Row = vGridMontos.ActiveRow
        If vGridMontos.MaxRows <= vGridMontos.ActiveRow Then
           vGridMontos.MaxRows = vGridMontos.MaxRows + 1
           vGridMontos.Row = vGridMontos.MaxRows
        End If
  End If 'Actualiza o Inserta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridMontos.MaxRows = vGridMontos.MaxRows + 1
    vGridMontos.InsertRows vGridMontos.ActiveRow, 1
    vGridMontos.Row = vGridMontos.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

    vGridMontos.Row = vGridMontos.ActiveRow
    vGridMontos.Col = 1

 If vGridMontos.Text = "" Then
     vGridMontos.DeleteRows vGridMontos.ActiveRow, 1
     vGridMontos.MaxRows = vGridMontos.MaxRows - 1
     If vGridMontos.MaxRows = 0 Then vGridMontos.MaxRows = 1
     Exit Sub
 End If
     
       i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
       If i = vbYes Then
                    
       strSQL = "delete afi_cd_actividades_rangos "
       strSQL = strSQL & "where cod_actividad = " & lswActividades.SelectedItem.Text
       vGridMontos.Col = 1
       strSQL = strSQL & "and cod_monto = vGridMontos.Text"
       Call ConectionExecute(strSQL)
                     
                      strSQL = vGridMontos.Text
                      vGridMontos.Col = 2
                      'Call Bitacora("Elimina", "Director: " & vGridMontos.Text & ")
                     
                      vGridMontos.DeleteRows vGridMontos.ActiveRow, 1
                      vGridMontos.MaxRows = vGridMontos.MaxRows - 1
                      If vGridMontos.MaxRows = 0 Then vGridMontos.MaxRows = 1
                     
          End If
          rs.Close
   End If


Exit Sub

vError:
  MsgBox Err.Description, vbCritical
 
End Sub

Private Function fxGuardarMontos() As Long

Dim strSQL As String, rs As New ADODB.Recordset

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError
Dim vTipo As String

fxGuardarMontos = 0
vGridMontos.Row = vGridMontos.ActiveRow
vGridMontos.Col = 1

If vGridMontos.Text = "" Then

    
    strSQL = "insert afi_cd_actividades_rangos(cod_actividad,monto,minimo,maximo,cod_monto) values(" & lswActividades.SelectedItem.Text & ","
    vGridMontos.Col = 2
    strSQL = strSQL & IIf((vGridMontos.Text = ""), 0, Int(vGridMontos.Text)) & ","
    vGridMontos.Col = 3
    strSQL = strSQL & IIf((vGridMontos.Text = ""), 0, Int(vGridMontos.Text)) & ","
    vGridMontos.Col = 4
    strSQL = strSQL & IIf((vGridMontos.Text = ""), 0, Int(vGridMontos.Value)) & ","
    vGridMontos.Col = 1
    vGridMontos.Text = fxConsecutivoCodigoMonto
    strSQL = strSQL & " " & vGridMontos.Text & ")"
    
    Call ConectionExecute(strSQL)
    MsgBox "Información Ingresada", vbInformation, "Información"
    

'    strSQL = vGridMontos.Text
    
    vGrid.Col = 2
    vGridMontos.MaxRows = vGridMontos.MaxRows + 1
    'Call Bitacora("Registra", "Directores: " & vGrid.Text & " Ced: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    fxGuardarMontos = 1
   
   Else 'Actualizar
  
   vGrid.Col = 2
   
    vGridMontos.Col = 2
    strSQL = "update afi_cd_actividades_rangos set monto = " & IIf((vGridMontos.Text = ""), 0, Int(vGridMontos.Text)) & ",minimo ="
    vGridMontos.Col = 3
    strSQL = strSQL & IIf((vGridMontos.Text = ""), 0, Int(vGridMontos.Text)) & ",maximo = "
    vGridMontos.Col = 4
    strSQL = strSQL & IIf((vGridMontos.Text = ""), 0, Int(vGridMontos.Value)) & " "
    vGridMontos.Col = 1
    strSQL = strSQL & "where cod_monto = " & vGridMontos.Text
    
    Call ConectionExecute(strSQL)
    MsgBox "Información Actualizada", vbInformation, "Información"
    
    strSQL = vGrid.Text
    
    vGrid.Col = 2
    'Call Bitacora("Modifica", "Directores: " & vGrid.Text & " ID: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    
   End If

Exit Function
vError:
MsgBox Err.Description, vbCritical
fxGuardarMontos = 0

End Function

Private Function fxGuardarActividad() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vActividad As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError
Dim vTipo As String

fxGuardarActividad = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then
    
    vGrid.Col = 1
    vGrid.Text = fxConsecutivo
    strSQL = "insert afi_cd_actividades(cod_actividad,descripcion,cod_cuenta,tipo,fechaperiocidad,fechaliq,activa) " _
    & "values(" & vGrid.Text & ",'"
    vGrid.Col = 2
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & "','"
    vGrid.Col = 3
    strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "','"
    vGrid.Col = 4
      If vGrid.Text = "Trimestral" Then
           vActividad = "T"
       Else
           vActividad = "E"
      End If
    strSQL = strSQL & IIf((vGrid.Text = ""), "", vActividad) & "','"
    vGrid.Col = 5
    strSQL = strSQL & IIf((vGrid.Text = ""), "", Format(vGrid.Text, "yyyy-mm-dd")) & "','"
    vGrid.Col = 6
    strSQL = strSQL & IIf((vGrid.Text = ""), "", Format(vGrid.Text, "yyyy-mm-dd")) & "',"
    vGrid.Col = 7
    strSQL = strSQL & IIf((vGrid.Value = ""), 0, vGrid.Value) & ")"
    
    
    
    Call ConectionExecute(strSQL)
    
'    strSQL = vGrid.Text
    
    vGrid.Row = vGrid.MaxRows + 1
    'Call Bitacora("Registra", "Directores: " & vGrid.Text & " Ced: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    fxGuardarActividad = 1
   
   Else 'Actualizar
  
   vGrid.Col = 2
   
    strSQL = "update afi_cd_actividades set descripcion  = '" & IIf((vGrid.Text = ""), 0, vGrid.Text) & "',cod_cuenta ='"
    vGrid.Col = 3
    strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "',tipo = '"
    vGrid.Col = 4
      If vGrid.Text = "Trimestral" Then
         vActividad = "T"
       Else
         vActividad = "E"
      End If
    strSQL = strSQL & IIf((vGrid.Text = ""), "", vActividad) & "',fechaperiocidad = '"
    vGrid.Col = 5
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, Format(vGrid.Text, "yyyymmdd")) & "',fechaliq = '" & ""
    vGrid.Col = 6
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, Format(vGrid.Text, "yyyymmdd")) & "',activa = "
    vGrid.Col = 7
    strSQL = strSQL & IIf((vGrid.Value = ""), 0, vGrid.Value) & " "
    vGrid.Col = 1
    strSQL = strSQL & "where cod_actividad = " & vGrid.Text
    
    Call ConectionExecute(strSQL)
    
    strSQL = vGrid.Text
    
    vGrid.Col = 2
    'Call Bitacora("Modifica", "Directores: " & vGrid.Text & " ID: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    
   End If

Exit Function
vError:
MsgBox Err.Description, vbCritical
fxGuardarActividad = 0
End Function
