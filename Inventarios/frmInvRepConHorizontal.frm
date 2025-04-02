VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvRepConHorizontal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidado Horizontal de Existencias"
   ClientHeight    =   6996
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6996
   ScaleWidth      =   7740
   Begin VB.CheckBox chkLsw 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3252
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   7416
      _ExtentX        =   13081
      _ExtentY        =   5736
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bodega"
         Object.Width           =   9596
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   7452
      _Version        =   1245187
      _ExtentX        =   13144
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   5280
         TabIndex        =   4
         Top             =   360
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Appearance      =   14
         Picture         =   "frmInvRepConHorizontal.frx":0000
      End
      Begin MSComctlLib.ProgressBar prgBarY 
         Height          =   132
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   4812
         _ExtentX        =   8488
         _ExtentY        =   233
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prgBarX 
         Height          =   132
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   4812
         _ExtentX        =   8488
         _ExtentY        =   233
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1200
      TabIndex        =   7
      Top             =   5160
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   252
      Left            =   2760
      TabIndex        =   9
      Top             =   5160
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Articulos sin Movimientos y Existencias  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   4
   End
   Begin XtremeSuiteControls.CheckBox chkCostos 
      Height          =   252
      Left            =   2760
      TabIndex        =   10
      Top             =   5400
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Costos de Articulos al Corte "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   4
      Value           =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   732
      Left            =   2160
      TabIndex        =   11
      Top             =   240
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Informe de Existencia: Bodegas"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Al Corte "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bodegas Disponibles"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7452
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmInvRepConHorizontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnReporte_Click()
Call sbReporte
End Sub

Private Sub chkLsw_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkLsw.Value
Next i

End Sub

Private Sub sbReporte()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, x As Integer

x = 0

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then x = x + 1
Next i

If x > 10 Then
  MsgBox "Este proceso no soporta mas de 10 Bodegas Seleccionadas", vbExclamation
  Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

prgBarX.Visible = True
prgBarY.Visible = True

prgBarY.Max = x + 4

'Limpiar Tabla Temporal
strSQL = "delete PV_INVREPTEMPO"
Call ConectionExecute(strSQL)
prgBarY.Value = 1

'Cargar Todos Los Productos a la Tabla Temporal
If chkCostos.Value = vbUnchecked Then
    strSQL = "insert into pv_InvRepTempo(cod_producto,costo,B0001,B0002,B0003,B0004,B0005" _
           & ",B0006,B0007,B0008,B0009,B0010) (select cod_producto,costo_regular,0,0,0,0,0,0,0,0,0,0" _
           & " from pv_productos)"
Else
    strSQL = "insert into pv_InvRepTempo(cod_producto,costo,B0001,B0002,B0003,B0004,B0005" _
           & ",B0006,B0007,B0008,B0009,B0010) (select cod_producto,dbo.fxINVCostoMercaderia(cod_producto,'" _
           & Format(dtpCorte.Value, "yyyy/mm/dd") & "'),0,0,0,0,0,0,0,0,0,0" _
           & " from pv_productos)"
End If
Call ConectionExecute(strSQL)
prgBarY.Value = 2

x = 1
For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
     'Procesar inventario en proceso para cada bodega en la fecha de corte\
     Call sbInvInventarioProceso(dtpCorte.Value, lsw.ListItems.Item(i).Text, False, False)
     
     'Carga Existencia en temporal
     strSQL = "select cod_producto,(existencia_inicial + entradas - salidas) as Existencia" _
            & " From pv_inventario_proceso" _
            & " where usuario = '" & glogon.Usuario & "'"
     Call OpenRecordSet(rs, strSQL)
     
     prgBarX.Value = 1
     prgBarX.Max = rs.RecordCount + 2
     
     Do While Not rs.EOF
       strSQL = "update pv_InvRepTempo set B" & Format(x, "0000") & " = " & rs!Existencia _
              & " where cod_producto = '" & rs!cod_producto & "'"
       Call ConectionExecute(strSQL)
       prgBarX.Value = prgBarX.Value + 1
       rs.MoveNext
     Loop
     rs.Close
     
     prgBarY.Value = prgBarY.Value + 1
     x = x + 1
  End If
Next i


If chkTodos.Value = vbUnchecked Then
  prgBarY.Value = prgBarY.Value + 1
  'Limpiar Tabla con Articulos sin movimientos ni Existencias
  strSQL = "delete pv_InvRepTempo where (B0001 = 0 and B0002 = 0 and B0003 = 0" _
         & " and B0004 = 0 and B0005 = 0 and B0006 = 0 and B0007 = 0 and B0008 = 0" _
         & " and B0009 = 0 and B0010 = 0)"
  Call ConectionExecute(strSQL)
End If


'Mostrar Reporte Aqui
prgBarY.Value = prgBarY.Value + 1
With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Invertarios"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"

 .Formulas(3) = "fxTitulo = 'INVENTARIO CONSOLIDADO - HORIZONTAL'"
 .Formulas(4) = "fxSubTitulo = 'Fecha Corte ...: " & Format(dtpCorte.Value, "dd/mm/yyyy") & IIf(chkCostos.Value = vbChecked, "Costos al Corte", "Costos Actuales") & "'"
         
 'Titulos de Bodegas
 x = 0
 For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
     x = x + 1
     .Formulas(4 + x) = "fxB" & Format(x, "0000") & " = '" & lsw.ListItems.Item(i).Text & "'"
  End If
 Next i
         
 For i = (x + 1) To 10
     .Formulas(4 + i) = "fxB" & Format(i, "0000") & " = 'N/A'"
 Next i
         
 .ReportFileName = SIFGlobal.fxPathReportes("Inventario_Consolidado.rpt")

 .PrintReport
End With

Me.MousePointer = vbDefault

prgBarX.Visible = False
prgBarY.Visible = False

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

dtpCorte.Value = fxFechaServidor

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

strSQL = "select cod_bodega,descripcion from pv_Bodegas"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_bodega)
     itmX.SubItems(1) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close

End Sub
