VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_RepRsmBalance 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumen de Balances"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   5292
      _Version        =   1572864
      _ExtentX        =   9340
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
   Begin XtremeSuiteControls.ComboBox cboCentroCosto 
      Height          =   312
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   5292
      _Version        =   1572864
      _ExtentX        =   9340
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
   Begin XtremeSuiteControls.ComboBox cboReporte 
      Height          =   312
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   5292
      _Version        =   1572864
      _ExtentX        =   9340
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   7815
      _Version        =   1572864
      _ExtentX        =   13785
      _ExtentY        =   2143
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   5880
         TabIndex        =   5
         Top             =   360
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_RepRsmBalance.frx":0000
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4812
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Costo"
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
      Index           =   6
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Negocio"
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
      Index           =   5
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2235
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmCntX_RepRsmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBusca As Integer, vPaso As Boolean



Private Sub cboUnidad_Click()
Dim strSQL As String

If vPaso Then Exit Sub

If cboUnidad.Text = "[CONSOLIDADO]" Then
    strSQL = "select rtrim(cod_Centro_Costo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
           & " from CntX_Centro_Costos where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Else
    strSQL = "select rtrim(cod_Centro_Costo) as 'IdX',rtrim(descripcion) as 'ItmX'" _
           & " from CntX_Centro_Costos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_centro_costo in(select cod_centro_costo from CntX_Unidades_CC where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "')"
End If


Call sbCbo_Llena_New(cboCentroCosto, strSQL, True, True)

End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String, vSubTitulo As String
Dim vSQL As String
Dim vPeriodoDesc As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vPeriodoDesc = fxCntX_PeriodoDesc(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 vSubTitulo = "PERIODO: " & vPeriodoDesc & " [PENDIENTE]"
Else
 vSubTitulo = "PERIODO: " & vPeriodoDesc & " [CERRADO]"
End If


With frmContenedor.Crt
     .Reset
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "ProGrX: Contabilidad"
     .Formulas(0) = "Empresa='" & gCntX_Parametros.NombreEmpresa & "'"
     .Formulas(1) = "Fecha='Fecha .:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(2) = "Usuario='Usuario .:" & glogon.Usuario & "'"
     .Formulas(3) = "SubTitulo='" & vSubTitulo & "'"
     .Formulas(4) = "Titulo='" & UCase(cboReporte.Text) & "'"
     
     .Connect = glogon.ConectRPT
    
 
    strSQL = "{vCntX_Mov_Rsm_TCuenta_Columna.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Rsm_TCuenta_Columna.ANIO} = " & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Rsm_TCuenta_Columna.MES} = " & gCntX_Parametros.PeriodoMes
    
    If cboUnidad.Text <> "[CONSOLIDADO]" Then
       strSQL = strSQL & " AND {vCntX_Mov_Rsm_TCuenta_Columna.COD_UNIDAD} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'"
    End If
              
    If cboCentroCosto.Text <> "TODOS" Then
       strSQL = strSQL & " AND {vCntX_Mov_Rsm_TCuenta_Columna.COD_CENTRO_COSTO} = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'"
    End If
               
    Select Case cboReporte.Text
      Case "Informe General por Unidades"
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_Rsm_Balance_Unidades.rpt")
      
      Case "Informe Resultados por Unidades"
                vSQL = "exec spCntX_BalanceRsmAnterior '" & glogon.Usuario & "'," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes
                glogon.Conection.Execute vSQL
               
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_Rsm_Resultados_Unidades.rpt")
      
      Case "Informe Resultados por Centros de Costo"
                 vSQL = "exec spCntX_BalanceRsmAnterior '" & glogon.Usuario & "'," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes
                 glogon.Conection.Execute vSQL
                
                .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_Rsm_Resultados_Cc.rpt")
    End Select
    
    .SelectionFormula = strSQL

     lbl.Caption = ""
    .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture
vBusca = 1

vPaso = True
Call sbCntX_CargaCboUnidades(cboUnidad)
vPaso = False

Call cboUnidad_Click

cboReporte.Clear
cboReporte.AddItem "Informe General por Unidades"
cboReporte.AddItem "Informe Resultados por Unidades"
cboReporte.AddItem "Informe Resultados por Centros de Costo"
cboReporte.Text = "Informe General por Unidades"


End Sub


