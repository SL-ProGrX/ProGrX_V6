VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCR_ReportesCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Control del Módulo de Créditos"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   3026
   Icon            =   "frmCR_ReportesCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   8775
   Begin VB.CheckBox chkComite 
      Caption         =   "Agrupar por Comité"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   480
      Width           =   2052
   End
   Begin VB.CheckBox chkHistoria 
      Caption         =   "Comp. Ult.6 Meses"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   840
      Width           =   2052
   End
   Begin VB.CheckBox chkProvincias 
      Caption         =   "Agrupar por Provincias"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   840
      Width           =   2292
   End
   Begin VB.CheckBox chkResumen 
      Caption         =   "Reporte Resumen"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   480
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkFechas 
      Alignment       =   1  'Right Justify
      Caption         =   "Todas las fechas"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDe 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   36278
   End
   Begin VB.CheckBox chkTodos 
      Alignment       =   1  'Right Justify
      Caption         =   "Todos las Líneas"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   36278
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   8535
      Begin VB.OptionButton optReportes 
         Caption         =   "Boletas de Refundiciones"
         Height          =   255
         Index           =   21
         Left            =   4440
         TabIndex        =   29
         Top             =   2400
         Width           =   3855
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Movimientos Generales a Operaciones"
         Height          =   255
         Index           =   20
         Left            =   4440
         TabIndex        =   28
         Top             =   2040
         Width           =   3855
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Retenciones Activas"
         Height          =   255
         Index           =   19
         Left            =   4440
         TabIndex        =   27
         Top             =   1680
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Formalizaciones - Retenciones"
         Height          =   255
         Index           =   18
         Left            =   4440
         TabIndex        =   26
         Top             =   1320
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Formalizaciones - Crédito (Conta)"
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   25
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Formalizaciones vrs Tesoreria"
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   24
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Auxiliar de Creditos con Saldos"
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   23
         Top             =   240
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Retenciones (Abonos Ded.Pla.)"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Abonos A Cartera Creditos (Liq,Ref,Rec)"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   3975
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Formalizaciones y Desembolsos"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   3735
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Créditos - En Cobro Judicial"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Créditos Traspasados a fiadores"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   3735
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Listado Situacion Crediticia"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   3495
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Planilla vrs Envio-Cargado"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Creditos x Instituc. Largo Plazo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Creditos x Instituc. Corto Plazo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Operaciones Morosas (Ind/NoPlanilla)"
         Height          =   255
         Index           =   25
         Left            =   4440
         TabIndex        =   30
         Top             =   2760
         Width           =   3855
      End
   End
   Begin VB.Image imgImprime 
      Height          =   480
      Left            =   8040
      Picture         =   "frmCR_ReportesCredito.frx":030A
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Image imgBusqueda_Rapida 
      Height          =   252
      Index           =   0
      Left            =   8040
      Picture         =   "frmCR_ReportesCredito.frx":0AB6
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda Rápida"
      Top             =   120
      Width           =   252
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   5412
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frmCR_ReportesCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFechas_Click()
 If chkFechas.Value = 1 Then
  dtpDe.Enabled = False
  dtpHasta.Enabled = False
 Else
  dtpDe.Enabled = True
  dtpHasta.Enabled = True
 End If
End Sub

Private Sub chkTodos_Click()
 If chkTodos.Value = 1 Then
  txtCodigo.Enabled = False
  chkHistoria.Value = False
  imgBusqueda_Rapida(0).Enabled = False
  txtCodigo.Text = ""
  lblDescripcion.Caption = ""
 Else
  txtCodigo.Enabled = True
  chkHistoria.Value = False
  imgBusqueda_Rapida(0).Enabled = False
 End If
End Sub


Private Sub Form_Load()
dtpDe.Value = fxFechaServidor
dtpHasta.Value = dtpDe
Call optReportes_Click(0)

End Sub


Private Sub imgBusqueda_Rapida_Click(Index As Integer)

'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Convertir = "N"
gBusquedas.Resultado = ""
Select Case Index
  Case 0 'txtCodigo
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        frmBusquedas.Show vbModal
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          lblDescripcion.Caption = fxDescribeCodigo(Trim(txtCodigo))
          chkTodos.SetFocus
        End If
End Select

End Sub

Private Sub imgImprime_Click()
Dim strRuta As String, strSQL As String
Dim dateFecha As Date, i As Long

On Error GoTo vError
Me.MousePointer = vbHourglass

If chkHistoria.Value = 1 And Trim(txtCodigo) = "" Then
   MsgBox "Para emitir este reporte debe" & vbCrLf & "suministrar un Codigo de Crédito", vbExclamation, "Reporte de Comportamiento"
   Exit Sub
End If

imgImprime.BorderStyle = 1

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 
 .Connect = glogon.ConectRPT
 
Select Case True
    
  Case optReportes(0).Value 'Creditos activos x Instituciones Corto Plazo
       
       strSQL = "{REG_CREDITOS.ESTADO} = 'A' AND {REG_CREDITOS.SALDO} > 0"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And {REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) _
                 & "," & Month(dtpDe) & "," & Day(dtpDe) & ") to Date(" & Year(dtpHasta) _
                 & "," & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todos'"
       End If
       
       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN CREDITOS ACTIVOS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_ResumenCreditosActivosCLP.rpt")
       Else
          .Formulas(3) = "Titulo='DETALLE CREDITOS ACTIVOS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_DetalleCreditosActivosCLP.rpt")
       End If
  
       i = InputBox("Indique el Plazo Maximo para Definicion de Corto plazo se recomienda (12) : ", "Indique el Plazo Maximo")
       strSQL = strSQL & " AND {REG_CREDITOS.PLAZO} <= " & i
       .SelectionFormula = strSQL
  
  
  Case optReportes(1).Value 'Creditos activos x Instituciones Largo Plazo
       
       strSQL = "{REG_CREDITOS.ESTADO} = 'A' AND {REG_CREDITOS.SALDO} > 0"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And {REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) _
                 & "," & Month(dtpDe) & "," & Day(dtpDe) & ") to Date(" & Year(dtpHasta) _
                 & "," & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todos'"
       End If
       
       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN CREDITOS ACTIVOS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_ResumenCreditosActivosCLP.rpt")
       Else
          .Formulas(3) = "Titulo='DETALLE CREDITOS ACTIVOS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_DetalleCreditosActivosCLP.rpt")
       End If
  
       i = InputBox("Indique el Plazo Mínimo para Definicion de Largo plazo se recomienda (12) : ", "Indique el Plazo Maximo")
       strSQL = strSQL & " AND {REG_CREDITOS.PLAZO} > " & i
       .SelectionFormula = strSQL
  
  
  Case optReportes(2).Value 'Planilla vrs Envio-Cargado
       
       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN COMPARATIVO PLANILLA ENVIADO vrs RECIBIDO'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanillaComparativoRsm.rpt")
       Else
          .Formulas(3) = "Titulo='DETALLE COMPARATIVO PLANILLA ENVIADO vrs RECIBIDO'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanillaComparativo.rpt")
       End If
  
       i = InputBox("Indique la fecha de proceso a comparar ", "Fecha de Proceso - Planilla")
       strSQL = "{vCRDPlanillaRepComparativo.Proceso} = " & i
       .Formulas(2) = "SubTitulo='Fecha de Proceso ...: " & Format(i, "####-##") & "'"
       .SelectionFormula = strSQL
 
  Case optReportes(3).Value 'Listado Situacion Crediticia
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_ListadoSituacionCrediticia.rpt")
       
       .Formulas(2) = "SubTitulo='Al  " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
       strSQL = "{REG_CREDITOS.SALDO} > 0"
       .SelectionFormula = strSQL
  
    
  Case optReportes(6).Value 'Creditos traspaso a fiadores
    If chkHistoria.Value = 1 Then
       strSQL = "{REG_CREDITOS.PROCESO} = 'T' AND {REG_CREDITOS.ESTADO} = 'A'"
       strSQL = strSQL & " AND {REG_CREDITOS.SALDO} = 0"
               
       strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"

       dateFecha = DateAdd("m", -5, Format(fxFechaServidor, "dd/mm/yyyy"))

       strSQL = strSQL & " And {REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Year(dateFecha)
       strSQL = strSQL & "," & Month(dateFecha) & ",01)"
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenCrecimiento_CreditosTraspasados.rpt")
       .SelectionFormula = strSQL
    Else
       strSQL = "{REG_CREDITOS.PROCESO} = 'T' AND {REG_CREDITOS.ESTADO} = 'A'"
       strSQL = strSQL & " AND {REG_CREDITOS.SALDO} = 0"
       
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{REG_CREDITOS.FECHA_ENVIAPROCESO} in Date (" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date (" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todas'"
       End If
       .SelectionFormula = strSQL
       
       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN CREDITOS TRASPASADOS A FIADORES'"
          If chkProvincias.Value = 1 Then
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenResumenCreditosTraspasadosProvincias.rpt")
            .SubreportToChange = "Grafico"
            .SelectionFormula = "{CATALOGO.CODIGO}={Pm-REG_CREDITOS.CODIGO} AND " & strSQL
          Else
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenResumenCreditosTraspasados.rpt")
          End If
       Else
          .Formulas(3) = "Titulo='DETALLE CREDITOS TRASPASADOS A FIADORES'"
          If chkProvincias.Value = 1 Then
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenDetalleCreditosTraspasadosProvincias.rpt")
          Else
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenDetalleCreditosTraspasados.rpt")
          End If
       End If
    End If
    

  Case optReportes(8).Value 'Creditos en cobro judicial
    If chkHistoria.Value = 1 Then
       strSQL = "{REG_CREDITOS.PROCESO} = 'J'"
       strSQL = strSQL & " AND {REG_CREDITOS.SALDO} > 0"
               
       strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"

       dateFecha = DateAdd("m", -5, Format(fxFechaServidor, "dd/mm/yyyy"))

       strSQL = strSQL & " And {REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Year(dateFecha)
       strSQL = strSQL & "," & Month(dateFecha) & ",01)"
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenCrecimiento_CreditosCobroJudicial.rpt")
       .SelectionFormula = strSQL
    Else
       strSQL = "{REG_CREDITOS.PROCESO} = 'J'"
       strSQL = strSQL & " AND {REG_CREDITOS.SALDO} > 0"
       
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{REG_CREDITOS.FECHA_ENVIAPROCESO} in Date (" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date (" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todas'"
       End If
       .SelectionFormula = strSQL
       
       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN CREDITOS EN COBRO JUDICIAL'"
          If chkProvincias.Value = 1 Then
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenResumenCreditosCobroProvincias.rpt")
            .SubreportToChange = "Grafico"
            .SelectionFormula = "{CATALOGO.CODIGO}={Pm-REG_CREDITOS.CODIGO} AND " & strSQL
          Else
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenResumenCreditosCobro.rpt")
          End If
       Else
          .Formulas(3) = "Titulo='DETALLE CREDITOS EN COBRO JUDICIAL'"
          If chkProvincias.Value = 1 Then
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenDetalleCreditosCobroProvincias.rpt")
          Else
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenDetalleCreditosCobro.rpt")
          End If
       End If
    End If

  Case optReportes(9).Value 'Formalizaciones y Desembolsos
       strSQL = ""
       If chkTodos.Value = 0 Then
          strSQL = "{VFORMALIZADESEM.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          If Len(strSQL) > 0 Then strSQL = strSQL & " And "
          strSQL = strSQL & "{VFORMALIZADESEM.FECHAFORP} in Date(" & Year(dtpDe) _
                 & "," & Month(dtpDe) & "," & Day(dtpDe) & ") to Date(" & Year(dtpHasta) _
                 & "," & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Historico'"
       End If
       
      Select Case (MsgBox("- Si desea visualizar los créditos desembolsados presione / SI" & vbCrLf _
                & "- Si desea los créditos **NO** Desembolsados NO" & vbCrLf _
                & "- Si Desea ver Todos Presione CANCELAR", vbYesNoCancel))
      Case vbYes 'SI
          If Len(strSQL) > 0 Then strSQL = strSQL & " And "
          strSQL = strSQL & "ISNULL({VFORMALIZADESEM.NDOCUMENTO}) = FALSE "
          .Formulas(2) = .Formulas(2) & "' * DESEMBOLSADOS *'"
      
      Case vbNo 'NO
          If Len(strSQL) > 0 Then strSQL = strSQL & " And "
          strSQL = strSQL & "ISNULL({VFORMALIZADESEM.NDOCUMENTO}) = TRUE "
          .Formulas(2) = .Formulas(2) & "' * NO DESEMBOLSADOS *'"
      
      Case vbCancel 'CANCELA
          .Formulas(2) = .Formulas(2) & "'TODOS'"
      End Select
       
       .SelectionFormula = strSQL
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenFormalizaDesembolsos.rpt")


  Case optReportes(10).Value 'Abonos a Cartera de Creditos - Liq,Ref,Rec
'       .ReportFileName = SIFGlobal.fxPathReportes("AbonosCartera.rpt")
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenMovimientosOperaciones.rpt")
       strSQL = "{CREDITOS_DT.ESTADO} = 'A' AND {CREDITOS_DT.TCON} <> '1'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {CREDITOS_DT.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{CREDITOS_DT.FECHAS} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Historico'"
       End If
       .Formulas(3) = "Titulo='ABONOS A CARTERA DE CREDITOS POR CONCEPTOS (LIQ,REF,REC)'"
       .SelectionFormula = strSQL
       
       .SubreportToChange = "Morosidad"
        
        
       strSQL = "{MOROSIDAD.ESTADO} = 'C' AND {MOROSIDAD.TCON} <> '1'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {MOROSIDAD.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          strSQL = strSQL & " And "
          strSQL = strSQL & "{MOROSIDAD.FECULT} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
       End If
       .SelectionFormula = strSQL
  
  
  
  Case optReportes(13).Value 'Retenciones
'       .ReportFileName = SIFGlobal.fxPathReportes("AbonosCartera.rpt")
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenMovimientosOperaciones.rpt")
       strSQL = "{CREDITOS_DT.ESTADO} = 'A' AND {CREDITOS_DT.TCON} = '1'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {CREDITOS_DT.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{CREDITOS_DT.FECHAS} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Historico'"
       End If
       .Formulas(3) = "Titulo='ABONOS A CARTERA DE CREDITOS POR DEDUC. PLANILLA'"

       .SelectionFormula = strSQL

       .SubreportToChange = "Morosidad"
        
        
       strSQL = "{MOROSIDAD.ESTADO} = 'C' AND {MOROSIDAD.TCON} = '1'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {MOROSIDAD.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          strSQL = strSQL & " And "
          strSQL = strSQL & "{MOROSIDAD.FECULT} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       End If
       
       .SelectionFormula = strSQL

   
    Case optReportes(15).Value
       'strSQL = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'"
       strSQL = "({REG_CREDITOS.ESTADO} = 'A') and {CATALOGO.RETENCION} = 'N' AND {CATALOGO.POLIZA} = 'N'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todas'"
       End If
       .SelectionFormula = strSQL

       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN AUXILIAR DE CREDITOS CON SALDOS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenConciliacionSaldosResumen.rpt")
       Else
          .Formulas(3) = "Titulo='DETALLE AUXILIAR DE CREDITOS CON SALDOS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenConciliacionSaldos.rpt")
       End If
       .Formulas(4) = "Mascara='" & GLOBALES.gstrNiveles & "'"
      
      
      Case optReportes(16) 'Formalizaciones vrs Tesoreria
      
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenFormalvrsTeso.rpt")
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          
          strSQL = "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ")to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
          .SelectionFormula = strSQL
          
          strSQL = "{CHEQUES.OP} = {?Pm-REG_CREDITOS.ID_SOLICITUD}"
          .SubreportToChange = "Cheques"
          .SelectionFormula = strSQL
          
          
          
      Case optReportes(17) 'Formalizaciones Creditos (Conta)
      
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenFormalizaCreditos.rpt")
          
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          .Formulas(3) = "Titulo='FORMALIZACIONES DE CREDITOS'"
          .Connect = glogon.Conection.ConnectionString
          strSQL = "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ")to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
          strSQL = strSQL & " AND ({REG_CREDITOS.ESTADOSOL} = 'F' OR {REG_CREDITOS.ESTADOSOL} = 'N')"
          strSQL = strSQL & " AND {CATALOGO.RETENCION} = 'N' AND {CATALOGO.POLIZA} = 'N'"
          .SelectionFormula = strSQL
          
          
      Case optReportes(18) 'Formalizaciones Retenciones
      
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenFormalizaRetenciones.rpt")
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          .Formulas(3) = "Titulo='FORMALIZACIONES DE RETENCIONES'"
          
         If chkFechas.Value = vbChecked Then
            .Formulas(2) = "SubTitulo='HISTORICO DE RETENCIONES FORMALIZADAS'"
            strSQL = "{REG_CREDITOS.ESTADOSOL} = 'F'"
            strSQL = strSQL & " AND ({CATALOGO.RETENCION} = 'S' OR {CATALOGO.POLIZA} = 'S')"
         Else
            .Formulas(2) = "SubTitulo='Formalizadas del " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
            strSQL = "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
            strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
            strSQL = strSQL & ")to Date(" & Year(dtpHasta) & ","
            strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
            strSQL = strSQL & " AND {REG_CREDITOS.ESTADOSOL} = 'F' AND ({CATALOGO.RETENCION} = 'S' OR {CATALOGO.POLIZA} = 'S')"
         End If
          
         If chkTodos.Value = vbUnchecked Then
            strSQL = strSQL & " AND {CATALOGO.CODIGO}= '" & txtCodigo & "'"
         End If
          
          .SelectionFormula = strSQL
      
      
      Case optReportes(19) 'Retenciones Acticas
      
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenFormalizaRetenciones.rpt")
          .Formulas(3) = "Titulo='RETENCIONES ACTIVAS'"
          
         If chkFechas.Value = vbChecked Then
            .Formulas(2) = "SubTitulo='LISTADO DE OPERACIONES ACTIVAS'"
            strSQL = "{REG_CREDITOS.ESTADO} = 'A'"
            strSQL = strSQL & " AND ({CATALOGO.RETENCION} = 'S' OR {CATALOGO.POLIZA} = 'S')"
         Else
            .Formulas(2) = "SubTitulo='Formalizadas del " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
            strSQL = "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
            strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
            strSQL = strSQL & ")to Date(" & Year(dtpHasta) & ","
            strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
            strSQL = strSQL & " AND {REG_CREDITOS.ESTADO} = 'A' AND ({CATALOGO.RETENCION} = 'S' OR {CATALOGO.POLIZA} = 'S')"
         End If
         
         If chkTodos.Value = vbUnchecked Then
            strSQL = strSQL & " AND {CATALOGO.CODIGO}= '" & txtCodigo & "'"
         End If
          
          .SelectionFormula = strSQL
          
          
          
  Case optReportes(20).Value 'Movimientos Generales a Operaciones
       
       If chkResumen.Value = vbChecked Then
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenMovimientosOperacionesRsm.rpt")
       Else
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenMovimientosOperaciones.rpt")
       End If
       
       If chkTodos.Value = 0 Then
          strSQL = "{CREDITOS_DT.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          If Len(strSQL) > 0 Then strSQL = strSQL & " And "
          strSQL = strSQL & "{CREDITOS_DT.FECHAS} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Historico'"
       End If
       .Formulas(3) = "Titulo='MOVIMIENTOS GENERALES A OPERACIONES'"
       .SelectionFormula = strSQL
       
       .SubreportToChange = "Morosidad"
        
        
       strSQL = "{MOROSIDAD.ESTADO} = 'C'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {MOROSIDAD.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          strSQL = strSQL & " And "
          strSQL = strSQL & "{MOROSIDAD.FECULT} in Date(" & Format(dtpDe.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
       Else
       End If
       .SelectionFormula = strSQL
          
          
  Case optReportes(21).Value 'Boletas de Refundiciones
       strSQL = "{REG_CREDITOS.ESTADOSOL} = 'F'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todos'"
       End If
       .SelectionFormula = strSQL
       .Formulas(3) = "Titulo='BOLETA DE REFUNDICIONES'"
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenListadoRefundiciones.rpt")
          
          
          
          
    Case optReportes(25).Value 'Casos Morosos que no se les deduce por Planilla
       strSQL = "{REG_CREDITOS.ESTADO} = 'A' AND {REG_CREDITOS.IND_DEDUCE_PLANILLA} = 'N'"
       If chkTodos.Value = 0 Then
          strSQL = strSQL & " And {REG_CREDITOS.CODIGO} ='" & Trim(txtCodigo) & "'"
       End If
       
       If chkFechas.Value = 0 Then
          .Formulas(2) = "SubTitulo='Del  " & Format(dtpDe, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
          strSQL = strSQL & " And "
          strSQL = strSQL & "{REG_CREDITOS.FECHAFORP} in Date(" & Year(dtpDe) & ","
          strSQL = strSQL & Month(dtpDe) & "," & Day(dtpDe)
          strSQL = strSQL & ") to Date(" & Year(dtpHasta) & ","
          strSQL = strSQL & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
       Else
          .Formulas(2) = "SubTitulo='Todas'"
       End If
       .SelectionFormula = strSQL

       If chkResumen.Value = 1 Then
          .Formulas(3) = "Titulo='RESUMEN CREDITOS MOROSOS QUE NO SE DEDUCEN X PLANILLA'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenResumenCreditosNoPlanillaMora.rpt")
       Else
          .Formulas(3) = "Titulo='DETALLE OPERACIONES MOROSAS QUE NO SE DEDUCEN X PLANILLA'"
          .ReportFileName = SIFGlobal.fxPathReportes("Credito_GenDetalleCreditosNoPlanillaMora.rpt")
       End If
          
          
          
End Select
  .PrintReport
End With

imgImprime.BorderStyle = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub optReportes_Click(Index As Integer)
Dim i As Integer

On Error Resume Next
For i = 0 To 30
 optReportes(i).FontBold = False
 optReportes(i).ForeColor = vbBlack
Next i

 optReportes(Index).FontBold = True
 optReportes(Index).ForeColor = vbBlue

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Dim rs As New ADODB.Recordset

On Error GoTo vError

If KeyAscii = vbKeyReturn Then
    rs.Source = "select codigo,descripcion from catalogo where codigo = '" & txtCodigo.Text & "'"
    rs.ActiveConnection = glogon.Conection
    rs.CursorType = adOpenStatic
    rs.Open
    
    If rs.EOF And rs.BOF Then
     MsgBox "No se encontró el código digitado...", vbCritical
    Else
     lblDescripcion = rs!Descripcion
     txtCodigo = rs!Codigo
     chkTodos.SetFocus
    End If
    rs.Close
 End If
 

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

