VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCntX_RepBalanceComprobacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Balance de Comprobación"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   2143
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   615
         Left            =   7320
         TabIndex        =   4
         Top             =   360
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Picture         =   "frmCntX_RepBalanceComprobacion.frx":0000
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
         TabIndex        =   5
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10398
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
   Begin XtremeSuiteControls.GroupBox fraRangos 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Cuentas:"
      ForeColor       =   8421504
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtDesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtHasta 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDesdeDesc 
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Top             =   360
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHastaDesc 
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Top             =   720
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
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
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   732
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   360
         Width           =   732
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   315
      Left            =   4080
      TabIndex        =   15
      Top             =   3720
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
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
      Height          =   315
      Left            =   4080
      TabIndex        =   16
      Top             =   4080
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
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
   Begin XtremeSuiteControls.CheckBox chkMovCero 
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   4560
      Width           =   3615
      _Version        =   1441793
      _ExtentX        =   6371
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
   Begin XtremeSuiteControls.CheckBox chkRevisar 
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   4800
      Width           =   3615
      _Version        =   1441793
      _ExtentX        =   6371
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Revisar Balance para Corregir Inconsistencias?"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
      _Version        =   1441793
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance de Comprobación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1875
      TabIndex        =   2
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label lblX01 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Informe"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblX01 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmCntX_RepBalanceComprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBusca As Integer

Private Sub sbBalanceActual(Optional vDivisaExtranjera As Boolean = False, Optional vMixto As Boolean = False)
Dim strSQL As String, strTitulo As String
Dim vPeriodoDesc As String

On Error GoTo vError

vPeriodoDesc = fxCntX_PeriodoDesc_Informes(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 strTitulo = vPeriodoDesc & " ¦ " & cboTipo.Text & " [PENDIENTE]"
Else
 strTitulo = vPeriodoDesc & " ¦ " & cboTipo.Text
End If

If Not fxCntX_BalanceCuadrado(CLng(gCntX_Parametros.PeriodoAnio), CInt(gCntX_Parametros.PeriodoMes)) Then
  MsgBox "El Balance de Comprobación no se encuentra cuadrado (Procesada en Utilitarios\Restructuración de Cuentas" _
        & " para solucionar este problema)...", vbCritical
End If

Select Case cboTipo.ItemData(cboTipo.ListIndex)
   Case 1
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {CntX_Cuentas.CUENTA_MADRE} ='' AND {vCntX_Mov_Cuentas_General.ANIO} = " _
           & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes
   Case 2
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes
   Case 3
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {vCntX_Mov_Cuentas_General.COD_CUENTA} >= '" & fxCntX_CuentaFormato(False, txtDesde) & "'" _
           & " AND {vCntX_Mov_Cuentas_General.COD_CUENTA} <= '" & fxCntX_CuentaFormato(False, txtHasta) & "'"
    strTitulo = strTitulo & " Cuentas : " & txtDesde & " Hasta " & txtHasta
   
   Case 4
    strSQL = ""
    Do While Not IsNumeric(strSQL)
      strSQL = InputBox("Digite el Número de Nivel a Visualizar", "Mostrar resultados por Nivel")
    Loop
    gCntX_Parametros.vNivelRep = strSQL
    
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {CntX_Cuentas.NIVEL} = " & gCntX_Parametros.vNivelRep
    strTitulo = strTitulo & " Cuentas AL NIVEL : " & gCntX_Parametros.vNivelRep
   
   
End Select

If chkMovCero.Value = vbChecked Then
  strTitulo = strTitulo & " [Cta con Mov]"
  strSQL = strSQL & " AND (ABS({vCntX_Mov_Cuentas_General.SALDO_INICIAL}) + " _
          & " ABS({vCntX_Mov_Cuentas_General.TOTAL_DEBITOS}) + ABS({vCntX_Mov_Cuentas_General.TOTAL_CREDITOS}) > 0)"
End If

If Not vDivisaExtranjera Then
   If Not vMixto Then
        Call sbCntX_Reportes("BalanceComprobacion", strSQL, strTitulo)
   Else
        Call sbCntX_Reportes("BalanceComprobacionMixto", strSQL, strTitulo)
   End If
Else
    Call sbCntX_Reportes("BalanceComprobacionDivisaForanea", strSQL, strTitulo)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbLibros(Optional pTipoLibro As String = "Diario")
Dim strSQL As String, strTitulo As String
Dim vPeriodoDesc As String, pAceptaMov As Integer

On Error GoTo vError

vPeriodoDesc = fxCntX_PeriodoDesc_Informes(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 strTitulo = cboTipo.Text & " - Periodo: " & vPeriodoDesc & " [PENDIENTE]"
Else
 strTitulo = cboTipo.Text & " - Periodo: " & vPeriodoDesc & " [CERRADO]"
End If


If pTipoLibro = "Diario" Then
    pAceptaMov = 1
Else
    pAceptaMov = 0
End If

If Not fxCntX_BalanceCuadrado(CLng(gCntX_Parametros.PeriodoAnio), CInt(gCntX_Parametros.PeriodoMes)) Then
  MsgBox "El Balance de Comprobación no se encuentra cuadrado (Procesada en Utilitarios\Restructuración de Cuentas" _
        & " para solucionar este problema)...", vbCritical
End If

Select Case cboTipo.ItemData(cboTipo.ListIndex)
   Case 1
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta & " AND {CntX_Cuentas.ACEPTA_MOVIMIENTOS} = " & pAceptaMov _
           & " AND {CntX_Cuentas.CUENTA_MADRE} ='' AND {vCntX_Mov_Cuentas_General.ANIO} = " _
           & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes
   Case 2
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta & " AND {CntX_Cuentas.ACEPTA_MOVIMIENTOS} = " & pAceptaMov _
           & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes
   Case 3
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta & " AND {CntX_Cuentas.ACEPTA_MOVIMIENTOS} = " & pAceptaMov _
           & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {vCntX_Mov_Cuentas_General.COD_CUENTA} >= '" & fxCntX_CuentaFormato(False, txtDesde) & "'" _
           & " AND {vCntX_Mov_Cuentas_General.COD_CUENTA} <= '" & fxCntX_CuentaFormato(False, txtHasta) & "'"
    strTitulo = strTitulo & " Cuentas : " & txtDesde & " Hasta " & txtHasta

   Case 4
    strSQL = ""
    Do While Not IsNumeric(strSQL)
      strSQL = InputBox("Digite el Número de Nivel a Visualizar", "Mostrar resultados por Nivel")
    Loop
    gCntX_Parametros.vNivelRep = strSQL
    
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta & " AND {CntX_Cuentas.ACEPTA_MOVIMIENTOS} = " & pAceptaMov _
           & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {CntX_Cuentas.NIVEL} = " & gCntX_Parametros.vNivelRep
    strTitulo = strTitulo & " Cuentas AL NIVEL : " & gCntX_Parametros.vNivelRep
   

End Select



If chkMovCero.Value = vbChecked Then
  strTitulo = strTitulo & " [Cta con Mov]"
  If pTipoLibro = "Diario" Then
        strSQL = strSQL & " AND (" _
                & " ABS({vCntX_Mov_Cuentas_General.TOTAL_DEBITOS}) + ABS({vCntX_Mov_Cuentas_General.TOTAL_CREDITOS}) > 0)"
  Else
        strSQL = strSQL & " AND (ABS({vCntX_Mov_Cuentas_General.SALDO_INICIAL}) + " _
                & " ABS({vCntX_Mov_Cuentas_General.TOTAL_DEBITOS}) + ABS({vCntX_Mov_Cuentas_General.TOTAL_CREDITOS}) > 0)"
  
  End If
  
End If

If pTipoLibro = "Diario" Then
     Call sbCntX_Reportes("LibrosDiario", strSQL, vPeriodoDesc)
Else
     Call sbCntX_Reportes("LibrosMayor", strSQL, vPeriodoDesc)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbBalanceActualUnidad(Optional vDivisaExtranjera As Boolean = False, Optional vMixto As Boolean = False)
Dim strSQL As String, strTitulo As String
Dim vUnidad As String
Dim vPeriodoDesc As String

On Error GoTo vError

vPeriodoDesc = fxCntX_PeriodoDesc_Informes(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 strTitulo = vPeriodoDesc & " ¦ " & cboTipo.Text & " [PENDIENTE]"
Else
 strTitulo = vPeriodoDesc & " ¦ " & cboTipo.Text
End If


vUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

If Not fxCntX_BalanceCuadrado(CLng(gCntX_Parametros.PeriodoAnio), CInt(gCntX_Parametros.PeriodoMes), vUnidad) Then
  MsgBox "El Balance de Comprobación no se encuentra cuadrado (Procesada en Utilitarios\Restructuración de Cuentas" _
        & " para solucionar este problema)...", vbCritical
End If

Select Case cboTipo.ItemData(cboTipo.ListIndex)
   Case 1
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {CntX_Cuentas.CUENTA_MADRE} ='' AND {vCntX_Mov_Cuentas_Unidad.ANIO} = " _
           & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_Unidad.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {vCntX_Mov_Cuentas_Unidad.cod_unidad} = '" & vUnidad & "'"
   Case 2
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Cuentas_Unidad.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_Unidad.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {vCntX_Mov_Cuentas_Unidad.cod_unidad} = '" & vUnidad & "'"
   Case 3
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Cuentas_Unidad.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_Unidad.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {vCntX_Mov_Cuentas_Unidad.cod_unidad} = '" & vUnidad & "'" _
           & " AND {vCntX_Mov_Cuentas_Unidad.COD_CUENTA} >= '" & fxCntX_CuentaFormato(False, txtDesde) & "'" _
           & " AND {vCntX_Mov_Cuentas_Unidad.COD_CUENTA} <= '" & fxCntX_CuentaFormato(False, txtHasta) & "'"
    strTitulo = strTitulo & " Cuentas : " & txtDesde & " Hasta " & txtHasta
   
   Case 4
    strSQL = ""
    Do While Not IsNumeric(strSQL)
      strSQL = InputBox("Digite el Número de Nivel a Visualizar", "Mostrar resultados por Nivel")
    Loop
    gCntX_Parametros.vNivelRep = strSQL
    
    strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
           & " AND {vCntX_Mov_Cuentas_Unidad.ANIO} = " & gCntX_Parametros.PeriodoAnio _
           & " AND {vCntX_Mov_Cuentas_Unidad.MES} = " & gCntX_Parametros.PeriodoMes _
           & " AND {vCntX_Mov_Cuentas_Unidad.cod_unidad} = '" & vUnidad & "'" _
           & " AND {CntX_Cuentas.NIVEL} = " & gCntX_Parametros.vNivelRep
    strTitulo = strTitulo & " Cuentas AL NIVEL : " & gCntX_Parametros.vNivelRep
   
   
End Select

If chkMovCero.Value = vbChecked Then
  strTitulo = strTitulo & " [Cta con Mov]"
  strSQL = strSQL & " AND (ABS({vCntX_Mov_Cuentas_Unidad.SALDO_INICIAL}) + " _
          & " ABS({vCntX_Mov_Cuentas_Unidad.TOTAL_DEBITOS}) + ABS({vCntX_Mov_Cuentas_Unidad.TOTAL_CREDITOS}) > 0)"
End If

If Not vDivisaExtranjera Then
   If Not vMixto Then
        Call sbCntX_Reportes("BalanceComprobacionU", strSQL, strTitulo)
   Else
        Call sbCntX_Reportes("BalanceComprobacionUMixto", strSQL, strTitulo)
   End If
Else
    Call sbCntX_Reportes("BalanceComprobacionUDivisaForanea", strSQL, strTitulo)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbBalancePreLiminar(Optional vUnidad As String = "0x0")
Dim strSQL As String, strTitulo As String
Dim vPeriodoDesc As String

On Error GoTo vError

vPeriodoDesc = fxCntX_PeriodoDesc_Informes(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)


If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 strTitulo = vPeriodoDesc & " ¦ " & cboTipo.Text & " (" & Trim(UCase(cboUnidad.Text)) & ")  [PENDIENTE]"
Else
 strTitulo = vPeriodoDesc & " ¦ " & cboTipo.Text & " (" & Trim(UCase(cboUnidad.Text)) & ")"
End If


Select Case cboTipo.ItemData(cboTipo.ListIndex)
   Case 1
    If vUnidad = "0x0" Then
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {CntX_Cuentas.CUENTA_MADRE} ='' AND {vCntX_Mov_Cuenta_Tmp.USUARIO} = '" _
               & glogon.Usuario & "'"
    Else
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {CntX_Cuentas.CUENTA_MADRE} ='' AND {CntX_Mov_Cuenta_Tmp.USUARIO} = '" _
               & glogon.Usuario & "'"
    End If
   
   Case 2
    If vUnidad = "0x0" Then
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {vCntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "'"
    Else
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {CntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "'"
    End If
   
   Case 3
    If vUnidad = "0x0" Then
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {vCntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "'" _
               & " AND {vCntX_Mov_Cuenta_Tmp.COD_CUENTA} >= '" & fxCntX_CuentaFormato(False, txtDesde) & "'" _
               & " AND {vCntX_Mov_Cuenta_Tmp.COD_CUENTA} <= '" & fxCntX_CuentaFormato(False, txtHasta) & "'"
    Else
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {CntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "'" _
               & " AND {CntX_Mov_Cuenta_Tmp.COD_CUENTA} >= '" & fxCntX_CuentaFormato(False, txtDesde) & "'" _
               & " AND {CntX_Mov_Cuenta_Tmp.COD_CUENTA} <= '" & fxCntX_CuentaFormato(False, txtHasta) & "'"
    End If
    strTitulo = strTitulo & " Cuentas : " & txtDesde & " Hasta " & txtHasta
   
   
   Case 4
    strSQL = ""
    Do While Not IsNumeric(strSQL)
      strSQL = InputBox("Digite el Número de Nivel a Visualizar", "Mostrar resultados por Nivel")
    Loop
    gCntX_Parametros.vNivelRep = strSQL
    
    If vUnidad = "0x0" Then
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {vCntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "'" _
               & " AND {CntX_Cuentas.NIVEL} = " & gCntX_Parametros.vNivelRep
    Else
        strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
               & " AND {CntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "'" _
               & " AND {CntX_Cuentas.NIVEL} = " & gCntX_Parametros.vNivelRep
    End If
    strTitulo = strTitulo & " Cuentas AL NIVEL : " & gCntX_Parametros.vNivelRep
   
   
End Select

If chkMovCero.Value = vbChecked Then
  strTitulo = strTitulo & " [Cta con Mov]"
  If vUnidad = "0x0" Then
    strSQL = strSQL & " AND (ABS({vCntX_Mov_Cuenta_Tmp.SI}) + " _
            & " ABS({vCntX_Mov_Cuenta_Tmp.TD}) + ABS({vCntX_Mov_Cuenta_Tmp.TC}) > 0)"
  Else
    strSQL = strSQL & " AND (ABS({CntX_Mov_Cuenta_Tmp.SI}) + " _
            & " ABS({CntX_Mov_Cuenta_Tmp.TD}) + ABS({CntX_Mov_Cuenta_Tmp.TC}) > 0)"
  End If
End If


If vUnidad = "0x0" Then
    Call sbCntX_Reportes("BALANCECOMPROBACIONPRECON", strSQL, strTitulo)
Else
    Call sbCntX_Reportes("BALANCECOMPROBACIONPRE", strSQL, strTitulo)
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub cbo_Click()


If cbo.ListCount > 0 Then
 If cbo.Text = "Libro de Diario" Or cbo.Text = "Libro de Mayor" Then
    cboTipo.Text = "Detallado"
    cboTipo.Enabled = False
 Else
    cboTipo.Enabled = True
 End If
End If

End Sub

Private Sub cmdReporte_Click()
Dim i As Integer, frmX As Form


If chkRevisar.Value = vbChecked Then
    i = MsgBox("Esta seguro que desea Restructurar los movimientos de este periodo...", vbYesNo)
    If i = vbYes Then
       Set frmX = frmCntX_Procesos
       Call sbCntX_RestructuraMovimientosRSM(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes, frmX)
    End If
End If


lbl.Visible = True
Select Case cbo.ItemData(cbo.ListIndex)
  Case 1 'Actual
    If Mid(cboUnidad.Text, 1, 2) = "[C" Then
        Call sbBalanceActual(False, False)
    Else
        Call sbBalanceActualUnidad(False, False)
    End If
  
  Case 2 'Actual Mixto con Divisas Foraneas
    If Mid(cboUnidad.Text, 1, 2) = "[C" Then
        Call sbBalanceActual(False, True)
    Else
        Call sbBalanceActualUnidad(False, True)
    End If
    
  
  Case 3 'Preliminar Acumulativo
    If Mid(cboUnidad.Text, 1, 2) = "[C" Then
        lbl.Caption = "Procesando Balance Preliminar [...Espere!]"
        lbl.Refresh
        Call sbCntX_Preliminar_Montar("A")
        Call sbBalancePreLiminar
    Else
        lbl.Caption = "Procesando Balance Preliminar [...Espere!]"
        lbl.Refresh
        Call sbCntX_Preliminar_Montar("A", cboUnidad.ItemData(cboUnidad.ListIndex))
        Call sbBalancePreLiminar(cboUnidad.ItemData(cboUnidad.ListIndex))
    End If

  Case 4 'Balance con Divisa Extranjera
    If Mid(cboUnidad.Text, 1, 2) = "[C" Then
        Call sbBalanceActual(True)
    Else
        Call sbBalanceActualUnidad(True)
    End If
    
  Case 5 'Libros Diarios
    Call sbLibros("Diario")
    
  Case 6 'LibrosMayor
    Call sbLibros("Mayor")
    
End Select

lbl.Caption = ""


End Sub


Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
vBusca = 1
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture
lbl.Caption = ""

Call sbCntX_CargaCboUnidades(cboUnidad)

cbo.AddItem "Balance Actual"
cbo.ItemData(cbo.ListCount - 1) = CStr(1)
cbo.AddItem "Balance Mixto con Divisas Foráneas"
cbo.ItemData(cbo.ListCount - 1) = CStr(2)
cbo.AddItem "Balance Preliminar Acumulativo"
cbo.ItemData(cbo.ListCount - 1) = CStr(3)
cbo.AddItem "Balance Divisa Extranjera"
cbo.ItemData(cbo.ListCount - 1) = CStr(4)

cbo.AddItem "Libro de Diario"
cbo.ItemData(cbo.ListCount - 1) = CStr(5)
cbo.AddItem "Libro de Mayor"
cbo.ItemData(cbo.ListCount - 1) = CStr(6)


cbo.Text = "Balance Actual"


cboTipo.AddItem "Resumido"
cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(1)

cboTipo.AddItem "Detallado"
cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(2)

cboTipo.AddItem "Por Rango de Ctas"
cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(3)

cboTipo.AddItem "Por Nivel de Cta"
cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(4)


cboTipo.Text = "Resumido"


'Si el Periodo Esta Cerrado (Bloquear Preliminares)
If Not fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
   chkPreliminar.Value = xtpUnchecked
   chkPreliminar.Enabled = False
Else
   chkPreliminar.Value = xtpChecked
   chkPreliminar.Enabled = True
End If



End Sub

Private Sub sbConsultas()
Select Case vBusca
  Case 1 'Desde
     
     frmCntX_ConsultaCuentas.Show vbModal
     txtDesde = gCuenta
  
  Case 2 'hasta

     frmCntX_ConsultaCuentas.Show vbModal
     txtHasta = gCuenta

End Select

End Sub


Private Sub txtDesde_GotFocus()
vBusca = 1
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtHasta.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub txtDesde_LostFocus()
 txtDesdeDesc.Text = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtDesde.Text))
 txtDesde.Text = fxCntX_CuentaFormato(True, txtDesde.Text)
End Sub

Private Sub txtHasta_GotFocus()
vBusca = 2
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub txtHasta_LostFocus()
 txtHastaDesc.Text = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtHasta.Text))
 txtHasta.Text = fxCntX_CuentaFormato(True, txtHasta.Text)
End Sub


