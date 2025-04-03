VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_Actividades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descripción de Actividades"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11985
   Icon            =   "FrmAF_CD_Actividades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   635
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
      BorderStyle     =   1
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
            Picture         =   "FrmAF_CD_Actividades.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":359C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":36AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_Actividades.frx":37D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Actividades"
      TabPicture(0)   =   "FrmAF_CD_Actividades.frx":38FD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Definición de Montos"
      TabPicture(1)   =   "FrmAF_CD_Actividades.frx":A15F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "lblActividad"
      Tab(1).Control(2)=   "imgRefrescar"
      Tab(1).Control(3)=   "vGridMontos"
      Tab(1).Control(4)=   "lswActividades"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Consulta"
      TabPicture(2)   =   "FrmAF_CD_Actividades.frx":109C1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dtpInicio"
      Tab(2).Control(1)=   "dtpCorte"
      Tab(2).Control(2)=   "lswConsulta"
      Tab(2).Control(3)=   "tlbConsulta"
      Tab(2).Control(4)=   "Line1"
      Tab(2).Control(5)=   "Label2(1)"
      Tab(2).Control(6)=   "Label2(0)"
      Tab(2).ControlCount=   7
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   -73800
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   168361987
         CurrentDate     =   40252
      End
      Begin MSComctlLib.ListView lswActividades 
         Height          =   4590
         Left            =   -74880
         TabIndex        =   1
         Top             =   720
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   8096
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Actividad"
            Object.Width           =   8290
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4800
         Left            =   285
         TabIndex        =   3
         ToolTipText     =   "Para seleccionar la cuenta contable (F4)"
         Top             =   480
         Width           =   11250
         _Version        =   524288
         _ExtentX        =   19844
         _ExtentY        =   8467
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         ScrollBars      =   2
         SpreadDesigner  =   "FrmAF_CD_Actividades.frx":17223
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridMontos 
         Height          =   4545
         Left            =   -69285
         TabIndex        =   4
         Top             =   735
         Width           =   5715
         _Version        =   524288
         _ExtentX        =   10081
         _ExtentY        =   8017
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "FrmAF_CD_Actividades.frx":17979
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   -73800
         TabIndex        =   9
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   168361987
         CurrentDate     =   40252
      End
      Begin MSComctlLib.ListView lswConsulta 
         Height          =   3870
         Left            =   -74760
         TabIndex        =   10
         Top             =   1440
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   6826
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Actividad"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comité"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Operación"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbConsulta 
         Height          =   360
         Left            =   -72360
         TabIndex        =   12
         Top             =   840
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         ButtonWidth     =   1958
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Consulta"
               Key             =   "Consulta"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Reporte"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -63720
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Image imgRefrescar 
         Height          =   240
         Left            =   -69840
         Picture         =   "FrmAF_CD_Actividades.frx":1800D
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblActividad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -69315
         TabIndex        =   2
         Top             =   360
         Width           =   5745
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actividades disponibles > Seleccione"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   5385
      End
   End
End
Attribute VB_Name = "frmAF_CD_Actividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vGuarda As Boolean

Dim Nuevo As Boolean
Dim Filas As Integer, Columnas As Integer, Inc As Integer, Can As Integer
Dim itmX As ListItem
Dim vCodigo As Integer
Dim vActivo As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Function fxConsecutivoCodigoMonto()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select coalesce(max(cod_monto),0) + 1 as Ultimo from afi_cd_actividades_rangos"
rs.Open strSQL, glogon.Conection, adOpenStatic
   fxConsecutivoCodigoMonto = rs!ultimo
rs.Close
End Function

Private Function fxConsecutivo()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select coalesce(max(cod_actividad),0) + 1 as Ultimo from AFI_CD_ACTIVIDADES"
rs.Open strSQL, glogon.Conection, adOpenStatic
   fxConsecutivo = rs!ultimo
rs.Close

End Function

Private Sub sbConsultaActividades()
Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim itmX As ListItem

lswConsulta.ListItems.Clear

strSQL = "select Act.descripcion,S.descripcion as Comite,C.monto,C.noperacion,C.registro_usuario,C.registro_fecha,C.estado " _
         & "from afi_cd_cuentas C inner join afi_cd_cuentas_actividades Ca on C.noperacion = Ca.noperacion " _
         & "inner join afi_cd_actividades Act on Act.cod_actividad = Ca.cod_actividad " _
         & "inner join afi_cd_comites S on C.Cod_comite = S.cod_comite " _
         & "where C.registro_fecha between '" & Format(dtpInicio.Value, "yyyymmdd 00:00:00") & "' " _
         & "and '" & Format(dtpCorte.Value, "yyyymmdd 23:59:59") & "'"
         
         rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
       Set itmX = lswConsulta.ListItems.Add(, , IIf(IsNull(rs!Descripcion), "", rs!Descripcion))
                  itmX.SubItems(1) = Trim(IIf(IsNull(rs!comite), "", rs!comite))
                  itmX.SubItems(2) = Trim(IIf(IsNull(rs!Monto), "", rs!Monto))
                  itmX.SubItems(3) = Trim(IIf(IsNull(rs!Noperacion), "", rs!Noperacion))
                  itmX.SubItems(4) = Trim(IIf(IsNull(rs!REGISTRO_USUARIO), "", rs!REGISTRO_USUARIO))
                  itmX.SubItems(5) = Trim(IIf(IsNull(rs!REGISTRO_FECHA), "", rs!REGISTRO_FECHA))
                  itmX.SubItems(6) = Trim(IIf(IsNull(rs!Estado), "", rs!Estado))
     rs.MoveNext
 Loop
rs.Close

End Sub

Private Sub sbControlLiquidacion()

Dim strSQL As String
Dim rs As New ADODB.Recordset
 


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
 
 .ReportFileName = SIFGlobal.fxSIFPathReportes("vista_afi_cd_cuentasactivas")
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
rs.Open strSQL, glogon.Conection, adOpenStatic
        
With lswActividades.ListItems
  .Clear
  Do While Not rs.EOF
      Set itmX = .Add(, , rs!Cod_actividad)
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

rs.Open strSQL, glogon.Conection, adOpenStatic

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
                vGrid.Text = rs!Cod_actividad
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
                vGrid.Value = rs!activa
         
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
 
 .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_Cd_Actividades.rpt")
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
 
 .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_Cd_Actividades.rpt")
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

Private Sub Form_Activate()
 vModulo = 23
End Sub

Private Sub Form_Load()
 
 vModulo = 23
  
 vActivo = False
 ssTab.Tab = 0
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

 Call sbCargaGridLocal
 
 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value

End Sub

Private Sub imgRefrescar_Click()
 Call sblswActividadesividadesCarga
End Sub

Private Sub lswActividades_Click()
If lswActividades.ListItems.Count <= 0 Then Exit Sub

vCodigo = lswActividades.SelectedItem
lblActividad.Caption = Trim(lswActividades.SelectedItem.SubItems(1))

strSQL = " select cod_monto,monto,minimo,maximo from afi_cd_actividades_rangos where cod_actividad = " & vCodigo & ""
Call sbCargaGrid(vGridMontos, 4, strSQL)

End Sub

Private Sub OptAct_Click()
 Call sbCargaGridLocal
End Sub

Private Sub OptDes_Click()
 Call sbCargaGridLocal
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
 
If ssTab.Tab = 1 Then
 Call sblswActividadesividadesCarga
End If
 
End Sub

Private Sub vGridact_DblClick(ByVal Col As Long, ByVal Row As Long)
 Call sbCargaCuenta
End Sub


Private Sub vGridact_KeyDown(KeyCode As Integer, Shift As Integer)

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim Conse As Integer, Inc As Integer

strSQL = "select coalesce(max(codtipo),0) + 1 as Ultimo from afi_cd_periocidadactividades"
          rs.Open strSQL, glogon.Conection, adOpenStatic

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


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
  Case "REPORTE"
     Call sbReporte
  Case "ACTUALIZA"
     Call sbCambiaAno
End Select
End Sub

Private Sub tlbConsulta_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
  Case "CONSULTA"
     Call sbConsultaActividades
  Case "REPORTE"
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
      .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ControlDesembolsosActividades.rpt")
      .Formulas(0) = "fxTitulo= 'CONTROL DE DESEMBOLSOS SEGUN ACTIVIDADES'"
       strSQL = strSQL & "cdate({vAFI_CD_CuentasActividades.tesoreria_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
       strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
      .Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
      .Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
      .SelectionFormula = strSQL
      .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
      .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
      

      .PrintReport
   End With
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
                      glogon.Conection.Execute strSQL
                     
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
       glogon.Conection.Execute strSQL
                     
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
    
    glogon.Conection.Execute strSQL
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
    
    glogon.Conection.Execute strSQL
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
    
    
    
    glogon.Conection.Execute strSQL
    
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
    
    glogon.Conection.Execute strSQL
    
    strSQL = vGrid.Text
    
    vGrid.Col = 2
    'Call Bitacora("Modifica", "Directores: " & vGrid.Text & " ID: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    
   End If

Exit Function
vError:
MsgBox Err.Description, vbCritical
fxGuardarActividad = 0
End Function
