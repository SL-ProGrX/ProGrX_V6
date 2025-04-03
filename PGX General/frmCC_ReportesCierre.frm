VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCC_ReportesCierre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes de Cierre"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   960
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Informe"
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
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   7080
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6862
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6980
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6AAA
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6BD0
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6CDE
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6DEB
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":6F04
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":7032
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":713F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_ReportesCierre.frx":723F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCatalogo 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   7815
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   7455
         _Version        =   1441793
         _ExtentX        =   13150
         _ExtentY        =   8070
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todos"
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
      Begin MSComctlLib.Toolbar tlbDocumento 
         Height          =   312
         Left            =   6360
         TabIndex        =   2
         Top             =   0
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   556
         ButtonWidth     =   1799
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgArbol"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aceptar"
               Key             =   "Aceptar"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView ArbolReportes 
      Height          =   5160
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9102
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgArbol"
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFin 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Cierre por Procesos"
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
      Height          =   615
      Index           =   3
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCC_ReportesCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim bSeleccionado As Boolean


Private Sub sbCargaNodos()

Dim iNodosCheck As Integer
Dim xNode As Node, i As Integer

iNodosCheck = 0
With ArbolReportes
  .Nodes.Clear
  '
  Set xNode = .Nodes.Add(, , "Pr", "Reportes")
  xNode.Bold = True
  'Reportes de Afiliacion
  .Nodes.Add "Pr", tvwChild, "Afiliacion", "Afliación Y Renuncias"
  .Nodes.Add "Afiliacion", tvwChild, "Ing", "Ingresos y Renuncias", 2, 1
  .Nodes.Add "Afiliacion", tvwChild, "ReIng", "ReIngresos Por Zonas", 2, 1
  
  'Aporte obrero Patronal
  .Nodes.Add "Pr", tvwChild, "Aportes", "Aportes(Obrero/Patronal)"
  .Nodes.Add "Aportes", tvwChild, "SinAp", "Asociados que no posean aportes", 5, 1
  .Nodes.Add "Aportes", tvwChild, "Apa", "Asociados sólo poseen aporte patronal", 6, 1 ' (reafiliaciones que aún no tienen nuevos aportes) "
  .Nodes.Add "Aportes", tvwChild, "Apo", "Asociados sólo poseen aporte obrero", 7, 1 ' (asociados que están enviado su aporte a otra institución)"
  .Nodes.Add "Aportes", tvwChild, "UAfl", "Asociados que se han afiliado por única vez", 8, 1 '(están activos)"
 ' .Nodes.Add "Aportes", tvwChild, "LiqAfi", "Asociados que se liquidaron y en el mismo mes se afiliaron", 4, 1
  
  'Reportes de credito
  .Nodes.Add "Pr", tvwChild, "Creditos", "Reporte según estado del deudor, por fechas"
  .Nodes.Add "Creditos", tvwChild, "CrAfl-1", "Afiliación menor a 1 meses membresía (reingreso más de un mes de haberse liquidado)  ", 2, 1
  .Nodes.Add "Creditos", tvwChild, "CrAfi1", "Afiliación menor a 1 meses membresía, 1º Ingreso (nunca fueron asociados)", 3, 1
  .Nodes.Add "Creditos", tvwChild, "CrAfReing", "Afiliación menor a 1 meses membresía (reingreso inmediato a la renuncia)", 4, 1
  
  
  
  
End With

For i = 1 To ArbolReportes.Nodes.Count
   ArbolReportes.Nodes.Item(i).Expanded = True
Next i

Me.MousePointer = vbDefault

 
End Sub


Private Sub ArbolReportes_Click()
'vCodigos = ""
'chkTodos.Value = vbUnchecked
bSeleccionado = True

'Select Case ArbolReportes.SelectedItem.Key

'    Case "CrAfl-1"
''        fraCatalogo.Visible = True
''        imgImprime.Enabled = False
''        Call sbCargaCatalogo
'        LblFin.Visible = True
'        dtpFin.Visible = True
'    Case "CrAfi1"
''        fraCatalogo.Visible = True
''        imgImprime.Enabled = False
''        Call sbCargaCatalogo
'        LblFin.Visible = True
'        dtpFin.Visible = True
'    Case "CrAfReing"
''        fraCatalogo.Visible = True
''        imgImprime.Enabled = False
''        Call sbCargaCatalogo
'        LblFin.Visible = True
'        dtpFin.Visible = True
'    Case "ReIng"
'        LblFin.Visible = True
'        dtpFin.Visible = True
'    Case Else
'        LblFin.Visible = False
'        dtpFin.Visible = False
'
'End Select

End Sub

Private Sub btnReporte_Click()
Me.MousePointer = vbHourglass


On Error GoTo vError



If bSeleccionado Then
    With frmContenedor.Crt
     .Reset
     .WindowShowExportBtn = True
     .WindowShowGroupTree = True
     .WindowShowPrintBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "Reportes Crédito"
     
     .Connect = glogon.ConectRPT
     
     .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
     .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     
     Select Case ArbolReportes.SelectedItem.Key
        Case "LiqAfi"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 5
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreAportesIngresos.rpt")
             .Destination = crptToWindow
             
        Case "ReIng"
          .Formulas(3) = "SubTitulo = 'REINGRESOS POR ZONAS DEL :" & Format(dtpInicio.Value, "dd/mm/yyyy") & " AL " & Format(dtpFin.Value, "dd/mm/yyyy") & "'"
         strSQL = strSQL & "cdate({liquidacion.FECLIQ}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
                  strSQL = strSQL & ") to Date (" & Format(dtpFin.Value, "yyyy,mm,dd") & ")"
                 ' strSQL = strSQL & " and cdate({AFI_INGRESOS.FECHA_INGRESO}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
                 ' strSQL = strSQL & ") to Date (" & Format(dtpFin.Value, "yyyy,mm,dd") & ")"
                  strSQL = strSQL & " and {Socios.EstadoActual} = 'S'"
                  strSQL = strSQL & " and month({AFI_INGRESOS.FECHA_INGRESO}) = " & Month(dtpInicio) & " and year({AFI_INGRESOS.FECHA_INGRESO}) = " & Year(dtpInicio) & " "
            
              .SelectionFormula = strSQL
              
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_Reingresos_Zonas.rpt")
             .Destination = crptToWindow
        Case "UAfl"
              .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
              .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 11
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreAportesIngresos.rpt")
             .Destination = crptToWindow
             
             
        Case "Ing"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = 'Afiliaciones del cierre del mes de ' & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 9
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrIngreso.rpt")
             .Destination = crptToWindow
             .PrintReport
    
             .Formulas(3) = "fxTitulo = 'Liquidaciones del cierre del mes de ' & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 10
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrRenuncias.rpt")
             .Destination = crptToWindow
             
        Case "SinAp"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 6
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreAportes.rpt")
             .Destination = crptToWindow
             
         Case "Apa"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 7
'             .StoredProcParam(3) = "i"
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreAportes.rpt")
             .Destination = crptToWindow
             
        Case "Apo"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 8
'             .StoredProcParam(3) = "i"
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreAportes.rpt")
             .Destination = crptToWindow
             
        Case "CrAfl-1"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & ' de '  & '" & Format(dtpInicio, "dd/mm/yyyy") & "' & ' hasta ' & '" & Format(dtpFin, "dd/mm/yyyy") & "'"
             .StoredProcParam(2) = 2
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreCredito.rpt")
             .Destination = crptToWindow
        
        Case "CrAfi1"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & ' de '  & '" & Format(dtpInicio, "dd/mm/yyyy") & "' & ' hasta ' & '" & Format(dtpFin, "dd/mm/yyyy") & "'"
             .StoredProcParam(2) = 1
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreCredito.rpt")
             .Destination = crptToWindow
        
        Case "CrAfReing"
             .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
             .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & ' de '  & '" & Format(dtpInicio, "dd/mm/yyyy") & "' & ' hasta ' & '" & Format(dtpFin, "dd/mm/yyyy") & "'"
             .StoredProcParam(2) = 3
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrCierreCredito.rpt")
             .Destination = crptToWindow
             
        
     End Select
     
    ' .PrintReport
     .Action = 1
    End With
End If


Me.MousePointer = vbDefault

Exit Sub


vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkTodos_Click()

Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", lsw.Width - 1300
End With

Call sbCargaNodos

bSeleccionado = False
dtpInicio.Value = fxFechaServidor
dtpFin.Value = dtpInicio.Value

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCargaCatalogo()

lsw.ListItems.Clear

strSQL = "select codigo,descripcion from catalogo" _
       & " where activo = 1 and (retencion = 'N' and poliza = 'N') order by descripcion"
Call OpenRecordSet(rs, strSQL)
 
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Trim(rs!Codigo))
     itmX.SubItems(1) = rs!Descripcion
     rs.MoveNext
Loop
rs.Close
End Sub

Private Sub tlbDocumento_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim i As Integer
'vCodigos = "'"
'strCodigos = ""
'For i = 1 To lsw.ListItems.Count
'
'  If lsw.ListItems.Item(i).Checked = True Then
'    vCodigos = vCodigos & lsw.ListItems.Item(i).Text & "',"
'
'    strCodigos = strCodigos & "," & lsw.ListItems.Item(i).Text
'    vCodigos = vCodigos & "'"
'  End If
'Next i
'
'
'
'
'vCodigos = Mid(vCodigos, 1, Len(vCodigos) - 2)
'vCodigos = "[" & vCodigos & "]"
'strCodigos = Mid(strCodigos, 2, Len(strCodigos))
'fraCatalogo.Visible = False
'imgImprime.Enabled = True

End Sub
