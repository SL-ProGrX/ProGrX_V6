VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmConCierreParcialesCreditos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierres Parciales de Auxiliares"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   8295
      _Version        =   1441793
      _ExtentX        =   14631
      _ExtentY        =   2566
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   615
         Left            =   5280
         TabIndex        =   13
         Top             =   600
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmConCierreParcialesCreditos.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnAnalisis 
         Height          =   615
         Left            =   6720
         TabIndex        =   14
         Top             =   600
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Cubo"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmConCierreParcialesCreditos.frx":07BC
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   315
         Left            =   5280
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
      End
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6794
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Auxiliar de Créditos"
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
      Value           =   -1  'True
   End
   Begin VB.Timer TimerX 
      Interval        =   100
      Left            =   7800
      Top             =   1200
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   6720
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6794
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Proyección de Cartera"
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
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6794
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Producto Acumulado"
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
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6794
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Auxiliar de Fondos"
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
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6794
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Auxiliar de Fondos (Balance Contable)"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtMeses 
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
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
      Text            =   "12"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCorteUlt 
      Height          =   315
      Left            =   6720
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
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
      Text            =   "..."
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblMeses 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Meses"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblUltCorte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ultimo Corte Procesado..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha de Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierres Parciales de Auxiliares del Sistema"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmConCierreParcialesCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAnalisis_Click()

Select Case True
 Case OptX.Item(0).Value 'Cierre Cartera
    Call sbCierreParcial
 
 Case OptX.Item(1).Value 'Proyeccion
    Call sbProyeccion
 
 Case OptX.Item(2).Value 'Producto Acumulado
    Call sbProductoAcumulado
    
End Select

End Sub

Private Sub sbAuxiliar_Credito()
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Visible = True
lblStatus.Caption = "Constuyendo Informe [Espere!]"
lblStatus.Refresh

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Cierres parciales de Auxiliar de Crédito"
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "SUBTITULO='Cierre Parcial al Corte: " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
    .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
         
    .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
 
    If Mid(cboTipo.Text, 1, 1) = "D" Then
       .Formulas(4) = "Titulo='CREDITOS: AUXILIAR AL CORTE [DETALLE]'"
       .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCredito_Corte_Det.rpt")
       i = 0
    Else
       .Formulas(4) = "Titulo='CREDITOS: AUXILIAR AL CORTE [RESUMEN]'"
       .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCredito_Corte_Rsm.rpt")
       i = 1
    End If
    
    .StoredProcParam(0) = Format(dtpCorte.Value, "YYYY-MM-DD 23:59:59.999")
    .StoredProcParam(1) = glogon.Usuario
    .StoredProcParam(2) = i
    .SelectionFormula = ""
   
    .Action = 1
End With

lblStatus.Caption = ""

Me.MousePointer = vbDefault

Exit Sub

vError:

Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAuxiliar_Fondos()
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Visible = True
lblStatus.Caption = "Constuyendo Informe [Espere!]"
lblStatus.Refresh

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Cierres parciales de Auxiliar de Fondos"
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "SUBTITULO='Cierre Parcial al Corte: " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
    .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
         
    .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
 
    
    Select Case True
       Case OptX.Item(3).Value 'Saldos
            
            
            .Formulas(4) = "Titulo='FONDOS: AUXILIAR AL CORTE [" & UCase(cboTipo.Text) & "]'"
            
            If Mid(cboTipo.Text, 1, 1) = "R" Then
               .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondos_Corte_Rsm.rpt")
               i = 1
            Else
               .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondos_Corte_Det.rpt")
               i = 0
            End If


       Case OptX.Item(4).Value 'Completo
            .Formulas(4) = "Titulo='FONDOS: AUXILIAR COMPLETO [" & UCase(cboTipo.Text) & "]'"
            
            If Mid(cboTipo.Text, 1, 1) = "R" Then
               .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondos_Corte_Completo_Rsm.rpt")
               i = 1
            Else
               .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondos_Corte_Completo_Det.rpt")
               i = 0
            End If

    End Select
    
    .StoredProcParam(0) = Format(dtpCorte.Value, "YYYY-MM-DD 23:59:59.999")
    .StoredProcParam(1) = glogon.Usuario
    .StoredProcParam(2) = 1
    .StoredProcParam(3) = i
    .SelectionFormula = ""
            
            
    .Action = 1
End With


lblStatus.Caption = ""

Me.MousePointer = vbDefault

Exit Sub

vError:

Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnReporte_Click()

Select Case True
  Case OptX.Item(0).Value 'Creditos
   Call sbAuxiliar_Credito
  Case OptX.Item(3).Value, OptX.Item(4).Value 'Fondos
   Call sbAuxiliar_Fondos
End Select

End Sub

Private Sub Form_Load()

vModulo = 10

cboTipo.AddItem "Resumen"
cboTipo.AddItem "Detalle"
cboTipo.Text = "Resumen"

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

Call OptX_Click(0)

End Sub


Private Sub OptX_Click(Index As Integer)

btnAnalisis.Enabled = True
btnReporte.Enabled = True

lblTitulo.Caption = OptX.Item(Index).Caption

Select Case True
 Case OptX.Item(0).Value 'Cierre
    lblFecha.Caption = "Fecha de Corte"
    
    lblUltCorte.Visible = True
    txtCorteUlt.Visible = True
    
    txtMeses.Visible = False
    lblMeses.Visible = False
    
 Case OptX.Item(1).Value 'Proyección
    lblFecha.Caption = "Fecha de Inicio"
    
    lblUltCorte.Visible = False
    txtCorteUlt.Visible = False
    
    txtMeses.Visible = True
    lblMeses.Visible = True
    
    btnReporte.Enabled = False
 
 Case OptX.Item(2).Value 'Producto Acumulado
    lblFecha.Caption = "Fecha de Corte"
    
    lblUltCorte.Visible = True
    txtCorteUlt.Visible = True
    
    txtMeses.Visible = False
    lblMeses.Visible = False
   
    btnReporte.Enabled = False
 
 Case OptX.Item(3).Value, OptX.Item(4).Value   'Auxiliar de Fondos
    lblFecha.Caption = "Fecha de Corte"
    
    lblUltCorte.Visible = True
    txtCorteUlt.Visible = True
    
    txtMeses.Visible = False
    lblMeses.Visible = False
    
    btnAnalisis.Enabled = False
End Select


cboTipo.Visible = btnReporte.Enabled

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

dtpCorte.Value = fxFechaServidor
dtpCorte.MaxDate = dtpCorte.Value

txtCorteUlt.Text = "No Existen Cortes!"
txtCorteUlt.ToolTipText = ""

strSQL = "select * From CRD_CIERRE_PARCIAL_CORTES" _
       & " where Linea in(select MAX(linea) from CRD_CIERRE_PARCIAL_CORTES)"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

  txtCorteUlt.Text = Format(rs!Corte, "dd/mm/yyyy")
  txtCorteUlt.ToolTipText = "Registrado por:" & vbCrLf _
                          & "   Usuario..: " & rs!Registro_Usuario & vbCrLf _
                          & "   Fecha  ..: " & rs!Registro_Fecha & vbCrLf
  
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCierreParcial()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Procesando Posición de la Cartera al corte requerido. Este proceso puede durar varios minutos, Espere.!" & vbCrLf _
                  & " Se Procesaran..:" & vbCrLf _
                  & "  - Saldos al Corte" & vbCrLf _
                  & "  - Posición de la Morosidad" & vbCrLf _
                  & "  - Vistas de Analisis" & vbCrLf _
                  & "  - Actualización de Cubos para Consultas"
                  
lblStatus.Visible = True
lblStatus.Refresh

strSQL = "exec spCrdCierreParcial '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','" & glogon.Usuario & "',1"
Call ConectionExecute(strSQL)

lblStatus.Visible = False

Me.MousePointer = vbDefault
MsgBox "Cierre de Cartera procesado satisfactoriamente, verificar información en el cubo de cierre..!", vbInformation

Call sbInicializa

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbProductoAcumulado()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Procesando Producto Acumulado al Corte!"
                  
lblStatus.Visible = True
lblStatus.Refresh

strSQL = "exec spSIFAuxProdAcumPP '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call ConectionExecute(strSQL)

lblStatus.Visible = False

Me.MousePointer = vbDefault
MsgBox "Producto Acumulado procesado satisfactoriamente, verificar información en el cubo de: ProdAcumulado..!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbProyeccion()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Proyectando la recuperación de la Cartera a " & txtMeses.Text & " mes(es). Espere.!"

lblStatus.Visible = True
lblStatus.Refresh

strSQL = "exec spCrdProyectaCartera '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'," & txtMeses.Text
Call ConectionExecute(strSQL)

lblStatus.Visible = False

Me.MousePointer = vbDefault
MsgBox "Proyección de Cartera procesado satisfactoriamente, verificar información en el cubo de Proyección..!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


