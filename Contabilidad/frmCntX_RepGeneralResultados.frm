VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_RepGeneralResultados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Balance General y Estado de Resultados"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ComboBox cboReporte 
      Height          =   555
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9340
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   7575
      _Version        =   1572864
      _ExtentX        =   13356
      _ExtentY        =   2138
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   5880
         TabIndex        =   14
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
         Picture         =   "frmCntX_RepGeneralResultados.frx":0000
      End
      Begin VB.Label lbl 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   15
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.CheckBox chkMovCero 
      Height          =   252
      Left            =   3120
      TabIndex        =   9
      Top             =   3960
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No Mostrar Cuentas con Saldo en Cero"
      BackColor       =   16777215
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
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
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   5640
      TabIndex        =   6
      Top             =   3240
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.ComboBox cboNiveles 
      Height          =   312
      Left            =   5640
      TabIndex        =   7
      Top             =   3600
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.CheckBox chkTitulos 
      Height          =   252
      Left            =   3120
      TabIndex        =   10
      Top             =   4200
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Títulos de la estructura contable ?"
      BackColor       =   16777215
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkOrden 
      Height          =   252
      Left            =   3120
      TabIndex        =   11
      Top             =   4440
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Cuentas de Orden en el Balance ?"
      BackColor       =   16777215
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkPreliminar 
      Height          =   252
      Left            =   4080
      TabIndex        =   12
      Top             =   1560
      Width           =   3612
      _Version        =   1572864
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Calcular Estados Preliminares?"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkCuentasContables 
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   4800
      Width           =   4575
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Formato con Número de Cuentas ?"
      BackColor       =   16777215
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
      Alignment       =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Top             =   3600
      Width           =   732
   End
   Begin VB.Label lblX01 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   3240
      Width           =   852
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
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
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
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmCntX_RepGeneralResultados"
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

Private Sub sbBalanceGeneral()
Dim strSQL As String, vSubTitulo As String
Dim vAsientoCierre As Integer
Dim vPeriodoDesc As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vPeriodoDesc = fxCntX_PeriodoDesc(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 vSubTitulo = "PERIODO: " & vPeriodoDesc & " [PENDIENTE] [Nivel " & cboNiveles.Text & "]"
Else
 vSubTitulo = "PERIODO: " & vPeriodoDesc & " [CERRADO] [Nivel " & cboNiveles.Text & "]"
End If

vSubTitulo = vSubTitulo & "  Unidad: " & cboUnidad.Text & "   Centro Costo: " & cboCentroCosto.Text


lbl.Visible = True

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ProGrX: Contabilidad"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & gCntX_Parametros.NombreEmpresa & "'"
 .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
 .Formulas(3) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
 .Formulas(4) = "SubTitulo='" & vSubTitulo & "'"
 .Connect = glogon.ConectRPT

'Verifica si existe Asiento de Cierre Fiscal para no mostrar utilidad temporal
If fxCntX_MesFiscal(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
   vAsientoCierre = 1 'Si
Else
   vAsientoCierre = 0 'No
End If
 

If chkPreliminar.Value = vbUnchecked Then
            .Formulas(5) = "fxAsientoCierre = " & vAsientoCierre
            .Formulas(6) = "fxMuestraTitulo = " & chkTitulos.Value
            
        '-------------------------------------------------

        If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
               strSQL = "{vCntX_Mov_Cuentas_General.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                      & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
                      & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
               
               If chkMovCero.Value = vbChecked Then
                  strSQL = strSQL & " AND {vCntX_Mov_Cuentas_General.Saldo_Inicial} + {vCntX_Mov_Cuentas_General.Total_Debitos} + {vCntX_Mov_Cuentas_General.Total_Creditos} <> 0"
               End If
               
               
               If chkOrden.Value = vbChecked Then
                   .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneralCtaOrden.rpt")
               Else
                   If chkCuentasContables.Value = xtpChecked Then
                       .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneral_Ctas.rpt")
                   Else
                       .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneral.rpt")
                   End If
               End If
               
               .SelectionFormula = strSQL
        
        End If
        
        If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
            .Formulas(5) = "fxAsientoCierre = " & vAsientoCierre
            .Formulas(6) = "fxUnidad='" & cboUnidad.Text & "'"
               
               'Por unidades y Centros de Costos
                strSQL = "{vCntX_Mov_Cuentas_Unidad.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                      & " AND {vCntX_Mov_Cuentas_Unidad.cod_unidad} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                      & " AND {vCntX_Mov_Cuentas_Unidad.ANIO} = " & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_Unidad.MES} = " & gCntX_Parametros.PeriodoMes _
                      & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
               
               If chkMovCero.Value = vbChecked Then
                  strSQL = strSQL & " AND {vCntX_Mov_Cuentas_Unidad.Saldo_Inicial} + {vCntX_Mov_Cuentas_Unidad.Total_Debitos} + {vCntX_Mov_Cuentas_Unidad.Total_Creditos} <> 0"
               End If
               
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneralUnidad.rpt")
               .SelectionFormula = strSQL
        End If


        '-------------------------------------------------

        If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
            strSQL = "{vCntX_Mov_Cuentas_CentroCosto.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.cod_unidad} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.cod_Centro_Costo} = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.ANIO} = " & gCntX_Parametros.PeriodoAnio _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.MES} = " & gCntX_Parametros.PeriodoMes _
                  & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
           
           If chkMovCero.Value = vbChecked Then
              strSQL = strSQL & " AND ({vCntX_Mov_Cuentas_CentroCosto.Saldo_Inicial} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto.Total_Debitos} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto.Total_Creditos} <> 0)"
           End If
          
          .Formulas(5) = "fxAsientoCierre = " & vAsientoCierre
          .Formulas(6) = "fxUnidad='" & cboUnidad.Text & "     Centro de Costos: " & cboCentroCosto.Text & "'"
           
           .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneralUnidadCentroCosto.rpt")
           .SelectionFormula = strSQL
        End If


        If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
            
          .Formulas(5) = "fxAsientoCierre = " & vAsientoCierre
          .Formulas(6) = "fxUnidad='" & cboUnidad.Text & "     Centro de Costos: " & cboCentroCosto.Text & "'"
            
            strSQL = "{vCntX_Mov_Cuentas_CentroCosto_Rsm.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.cod_unidad} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.cod_Centro_Costo} = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.ANIO} = " & gCntX_Parametros.PeriodoAnio _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.MES} = " & gCntX_Parametros.PeriodoMes _
                  & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
           
           If chkMovCero.Value = vbChecked Then
              strSQL = strSQL & " AND ({vCntX_Mov_Cuentas_CentroCosto_Rsm.Saldo_Inicial} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto_Rsm.Total_Debitos} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto_Rsm.Total_Creditos} <> 0)"
           End If
           
           .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneralCentroCosto.rpt")
           .SelectionFormula = strSQL
        End If





Else
   'Calcular Preliminares
            .Formulas(5) = "fxMuestraTitulo = " & chkTitulos.Value
   
            lbl.Caption = "Procesando Balance Preliminar [...Espere!]"
            lbl.Refresh
            
            If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
                Call sbCntX_Preliminar_Montar("A")
            End If
            
            If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
                Call sbCntX_Preliminar_Montar("A", cboUnidad.ItemData(cboUnidad.ListIndex))
            End If
            
            If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
                Call sbCntX_Preliminar_Montar("A", cboUnidad.ItemData(cboUnidad.ListIndex), cboCentroCosto.ItemData(cboCentroCosto.ListIndex))
            End If
            
            If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
                Call sbCntX_Preliminar_Montar("A", "0x0", cboCentroCosto.ItemData(cboCentroCosto.ListIndex))
            End If
            
            
            strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                   & " AND {vCntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "' AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
               
            If chkMovCero.Value = vbChecked Then
             strSQL = strSQL & " AND (ABS({vCntX_Mov_Cuenta_Tmp.SI}) + " _
                     & " ABS({vCntX_Mov_Cuenta_Tmp.TD}) + ABS({vCntX_Mov_Cuenta_Tmp.TC}) > 0)"
            End If
            
            If chkCuentasContables.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneralPreliminar_Ctas.rpt")
            Else
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceGeneralPreliminar.rpt")
            End If
           
           .SelectionFormula = strSQL

End If 'Preliminar

lbl.Caption = ""


 .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbEstadoResultados()
Dim strSQL As String, vSubTitulo As String
Dim vPeriodoDesc As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vPeriodoDesc = fxCntX_PeriodoDesc(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 vSubTitulo = "PERIODO: " & vPeriodoDesc & " [PENDIENTE] [Nivel " & cboNiveles.Text & "]"
Else
 vSubTitulo = "PERIODO: " & vPeriodoDesc & " [CERRADO] [Nivel " & cboNiveles.Text & "]"
End If


vSubTitulo = vSubTitulo & "  Unidad: " & cboUnidad.Text & "   Centro Costo: " & cboCentroCosto.Text


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ProGrX: Contabilidad"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & gCntX_Parametros.NombreEmpresa & "'"
 .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
 .Formulas(3) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
 .Formulas(4) = "SubTitulo='" & vSubTitulo & "'"
 .Connect = glogon.ConectRPT

If chkPreliminar.Value = vbUnchecked Then

        .Formulas(5) = "fxMuestraTitulo = " & chkTitulos.Value

        If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
           strSQL = "{vCntX_Mov_Cuentas_General.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                  & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
                  & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
           
           If chkMovCero.Value = vbChecked Then
              strSQL = strSQL & " AND ({vCntX_Mov_Cuentas_General.Saldo_Inicial} <> 0 OR {vCntX_Mov_Cuentas_General.Total_Debitos} <> 0 OR {vCntX_Mov_Cuentas_General.Total_Creditos} <> 0)"
           End If
           
           If chkCuentasContables.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultados_Ctas.rpt")
           Else
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultados.rpt")
           End If
           
           .SelectionFormula = strSQL
        
        End If
        
        If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
            strSQL = "{vCntX_Mov_Cuentas_Unidad.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                  & " AND {vCntX_Mov_Cuentas_Unidad.cod_unidad} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_Unidad.ANIO} = " & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_Unidad.MES} = " & gCntX_Parametros.PeriodoMes _
                  & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
           
           If chkMovCero.Value = vbChecked Then
              strSQL = strSQL & " AND ({vCntX_Mov_Cuentas_Unidad.Saldo_Inicial} <> 0 OR {vCntX_Mov_Cuentas_Unidad.Total_Debitos} <> 0 OR {vCntX_Mov_Cuentas_Unidad.Total_Creditos} <> 0)"
           End If
          .Formulas(6) = "fxUnidad='" & cboUnidad.Text & "     Centro de Costos: " & cboCentroCosto.Text & "'"
           
           .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosUnidad.rpt")
           .SelectionFormula = strSQL
        End If



        If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
            strSQL = "{vCntX_Mov_Cuentas_CentroCosto.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.cod_unidad} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.cod_Centro_Costo} = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.ANIO} = " & gCntX_Parametros.PeriodoAnio _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto.MES} = " & gCntX_Parametros.PeriodoMes _
                  & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
           
           If chkMovCero.Value = vbChecked Then
              strSQL = strSQL & " AND ({vCntX_Mov_Cuentas_CentroCosto.Saldo_Inicial} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto.Total_Debitos} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto.Total_Creditos} <> 0)"
           End If
          .Formulas(6) = "fxUnidad='" & cboUnidad.Text & "     Centro de Costos: " & cboCentroCosto.Text & "'"
           
           .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosUnidadCentroCosto.rpt")
           .SelectionFormula = strSQL
        End If


        If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
            strSQL = "{vCntX_Mov_Cuentas_CentroCosto_Rsm.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.cod_unidad} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.cod_Centro_Costo} = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.ANIO} = " & gCntX_Parametros.PeriodoAnio _
                  & " AND {vCntX_Mov_Cuentas_CentroCosto_Rsm.MES} = " & gCntX_Parametros.PeriodoMes _
                  & " AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
           
           If chkMovCero.Value = vbChecked Then
              strSQL = strSQL & " AND ({vCntX_Mov_Cuentas_CentroCosto_Rsm.Saldo_Inicial} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto_Rsm.Total_Debitos} <> 0 OR {vCntX_Mov_Cuentas_CentroCosto_Rsm.Total_Creditos} <> 0)"
           End If
          .Formulas(6) = "fxUnidad='" & cboUnidad.Text & "     Centro de Costos: " & cboCentroCosto.Text & "'"
           
           .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosCentroCosto.rpt")
           .SelectionFormula = strSQL
        End If


Else
   'Calcular Preliminares
   
            .Formulas(5) = "fxMuestraTitulo = " & chkTitulos.Value
            
            lbl.Caption = "Procesando Balance Preliminar [...Espere!]"
            lbl.Refresh
            
            If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
                Call sbCntX_Preliminar_Montar("A")
            End If
            
            If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
                Call sbCntX_Preliminar_Montar("A", cboUnidad.ItemData(cboUnidad.ListIndex))
            End If
            
            If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
                Call sbCntX_Preliminar_Montar("A", cboUnidad.ItemData(cboUnidad.ListIndex), cboCentroCosto.ItemData(cboCentroCosto.ListIndex))
            End If
            
            If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
                Call sbCntX_Preliminar_Montar("A", "0x0", cboCentroCosto.ItemData(cboCentroCosto.ListIndex))
            End If
            
            
            
            strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                   & " AND {vCntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "' AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
               
            If chkMovCero.Value = vbChecked Then
             strSQL = strSQL & " AND (ABS({vCntX_Mov_Cuenta_Tmp.SI}) + " _
                     & " ABS({vCntX_Mov_Cuenta_Tmp.TD}) + ABS({vCntX_Mov_Cuenta_Tmp.TC}) > 0)"
            End If

            If chkCuentasContables.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosPreliminar_Ctas.rpt")
            Else
               .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosPreliminar.rpt")
            End If
           
           .SelectionFormula = strSQL

End If 'Preliminar

lbl.Caption = ""

 .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdReporte_Click()

Select Case cboReporte.Text
  Case "Balance General"
    Call sbBalanceGeneral
  Case "Estado de Resultados"
    Call sbEstadoResultados
End Select

End Sub


Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
vModulo = 20

vBusca = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vPaso = True
Call sbCntX_CargaCboUnidades(cboUnidad)
vPaso = False

Call cboUnidad_Click

cboReporte.Clear
cboReporte.AddItem "Balance General"
cboReporte.AddItem "Estado de Resultados"
cboReporte.Text = "Balance General"

cboTipo.AddItem "Resumen Agrupado"
cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(1)

cboTipo.AddItem "Detalle Agrupado"
cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(2)

cboTipo.Text = "Resumen Agrupado"


cboNiveles.Clear
cboNiveles.AddItem "1"
cboNiveles.AddItem "2"
cboNiveles.AddItem "3"
cboNiveles.AddItem "4"
cboNiveles.AddItem "5"
cboNiveles.AddItem "6"
cboNiveles.AddItem "7"
cboNiveles.AddItem "8"

cboNiveles.Text = "2"



'Si el Periodo Esta Cerrado (Bloquear Preliminares)
If Not fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
   chkPreliminar.Value = xtpUnchecked
   chkPreliminar.Enabled = False
Else
   chkPreliminar.Value = xtpChecked
   chkPreliminar.Enabled = True
End If

End Sub


