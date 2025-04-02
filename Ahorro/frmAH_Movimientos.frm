VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAH_Movimientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Movimientos a Patrimonio"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   15660
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   732
      Left            =   12240
      TabIndex        =   16
      Top             =   600
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmAH_Movimientos.frx":0000
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   15660
      _ExtentX        =   27623
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   14532
      _Version        =   524288
      _ExtentX        =   25633
      _ExtentY        =   8911
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
      MaxCols         =   482
      SpreadDesigner  =   "frmAH_Movimientos.frx":0A1E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   18000
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Movimientos.frx":1332
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Movimientos.frx":7B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Movimientos.frx":E3F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Movimientos.frx":14C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Movimientos.frx":14D8F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   556
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
      Height          =   312
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   4200
      TabIndex        =   12
      Top             =   600
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ComboBox cboTransaccion 
      Height          =   312
      Left            =   4200
      TabIndex        =   13
      Top             =   960
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   315
      Left            =   9600
      TabIndex        =   14
      Top             =   600
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
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
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   9600
      TabIndex        =   15
      Top             =   960
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   732
      Left            =   13440
      TabIndex        =   17
      ToolTipText     =   "Exportar"
      Top             =   600
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   1291
      _StockProps     =   79
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
      Picture         =   "frmAH_Movimientos.frx":14E8F
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Aporte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Transacción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No.Documento"
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
      Index           =   9
      Left            =   8160
      TabIndex        =   7
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Cédula"
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
      Index           =   10
      Left            =   8160
      TabIndex        =   6
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   240
      X2              =   3000
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmAH_Movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

If cbo.Text = "[TODOS]" Then
  vTipo = "'O','P','C','E','X'"
Else
  If cbo.Text = "Custodia" Then
    vTipo = "'X'"
  Else
    vTipo = "'" & Mid(cbo.Text, 1, 1) & "'"
  End If
End If

       
strSQL = "select D.*, Sec.DESCRIPCION as 'SectorDesc'" _
       & " from vSIF_CtrlDoc_Pat_Detalle D" _
       & " inner join Socios S on D.cedula = S.CEDULA" _
       & " left join AFI_SECTORES Sec on S.COD_SECTOR = Sec.COD_SECTOR" _
       & " Where D.Tipo_Aporte_Id in(" & vTipo & ")" _
       & " and D.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
              
              
If cboTransaccion.Text <> "TODOS" Then
   strSQL = strSQL & " and TCon = '" & cboTransaccion.ItemData(cboTransaccion.ListIndex) & "'"
End If
       
If Len(Trim(txtDocumento.Text)) > 0 Then
   strSQL = strSQL & " and NCon like '" & txtDocumento.Text & "%'"
End If

If Len(Trim(txtIdentificacion.Text)) > 0 Then
   strSQL = strSQL & " and Cedula like '" & txtIdentificacion.Text & "%'"
End If

       
vGrid.MaxRows = 0
vGrid.MaxCols = 14

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1

Do While Not rs.EOF
 vGrid.MaxRows = vGrid.MaxRows + 1
 vGrid.Row = vGrid.MaxRows
 For i = 1 To 14
   vGrid.col = i
   Select Case i
      Case 1 'Consecutivo
         vGrid.Text = CStr(rs!Id_seq)
      Case 2 'Tipo
         vGrid.Text = Trim(rs!Tipo_Aporte)
      Case 3 'Cedula
         vGrid.Text = Trim(rs!Cedula)
      Case 4 'Nombre
         vGrid.Text = Trim(rs!Nombre)
      Case 5 'Monto
         vGrid.Text = Format(rs!Monto)
      Case 6 'Fecha
         vGrid.Text = Trim(rs!fecha)
      Case 7 'Usuario
         vGrid.Text = Trim(rs!Usuario & "")
      Case 8 'Concepto
         vGrid.Text = Trim(rs!CONCEPTO)
      Case 9 'Tipo de documento
         vGrid.Text = Trim(rs!Tipo)
      Case 10 'No. Documento
         vGrid.Text = Trim(rs!nCon & "")
      Case 11 'Caja
         vGrid.Text = Trim(rs!Cod_Caja & "")
      Case 12 'Fecha Proceos
         vGrid.Text = Format(rs!FechaProc, "####-##")
      Case 13 'Institucion
         vGrid.Text = Trim(rs!Institucion & "")
      Case 14 'Sector
         vGrid.Text = Trim(rs!SectorDesc & "")
   End Select
 
 Next i
 
 PrgBar.Value = PrgBar.Value + 1
 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
MsgBox "Consulta Finalizada...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 14
    vHeaders.Headers(1) = " (Id) "
    vHeaders.Headers(2) = "Tipo"
    vHeaders.Headers(3) = "Cedula"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Monto"
    vHeaders.Headers(6) = "Fecha"
    vHeaders.Headers(7) = "Usuario"
    vHeaders.Headers(8) = "Concepto"
    vHeaders.Headers(9) = "Tipo Doc."
    vHeaders.Headers(10) = "Num. Doc."
    vHeaders.Headers(11) = "Caja"
    vHeaders.Headers(12) = "Proceso"
    vHeaders.Headers(13) = "Institución"
    vHeaders.Headers(14) = "Sector"
    
    
'Select Case ButtonMenu.Key
'  Case "Excel"
      Call sbSIFGridExportar(vGrid, vHeaders, "Patrimonio_Aportes_Mov")
'  Case "HTML"
'      Call sbSIFGridExportar(vGrid, vHeaders, "Patrimonio_Aportes_Mov", "HTML")
'End Select
End Sub

Private Sub Form_Load()

vModulo = 2

Call Formularios(Me)
Call RefrescaTags(Me)

Call sbInicializa

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 450
vGrid.Height = Me.Height - (vGrid.top + 500)

End Sub


Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

vGrid.MaxCols = 14
vGrid.MaxRows = 0

cbo.Clear
cbo.AddItem "Obrero"
cbo.AddItem "Patronal"
cbo.AddItem "Custodia"
cbo.AddItem "Capitalizado"
cbo.AddItem "[TODOS]"
cbo.Text = "[TODOS]"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

strSQL = "select rtrim(Tipo_Documento) as  'IdX',   rtrim(Descripcion) as 'ItmX'" _
       & " from sif_documentos Where Tipo_Documento in('ND','NC','RE','LIQ','RLIQ','PLA','ING','CAJA','CAJARE')" _
       & " order by descripcion"
Call sbCbo_Llena_New(cboTransaccion, strSQL, True, True)


Me.MousePointer = vbDefault

End Sub



