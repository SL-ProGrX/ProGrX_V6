VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_RepBalanceSituacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Balance de Situación"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox fraRangos 
      Height          =   1452
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   7812
      _Version        =   1572864
      _ExtentX        =   13779
      _ExtentY        =   2561
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtDesde 
         Height          =   312
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Height          =   312
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Height          =   312
         Left            =   3000
         TabIndex        =   18
         Top             =   360
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   550
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
         Height          =   312
         Left            =   3000
         TabIndex        =   19
         Top             =   720
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.ComboBox cboNiveles 
         Height          =   330
         Left            =   1320
         TabIndex        =   26
         Top             =   1080
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
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
         Index           =   7
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   855
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
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   615
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
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   5652
      _Version        =   1572864
      _ExtentX        =   9975
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
      TabIndex        =   2
      Top             =   1800
      Width           =   5652
      _Version        =   1572864
      _ExtentX        =   9975
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
      Left            =   480
      TabIndex        =   5
      Top             =   5880
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
         TabIndex        =   6
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
         Picture         =   "frmCntX_RepBalanceSituacion.frx":0000
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
         TabIndex        =   7
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.CheckBox chkAnalisiHV 
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   5160
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Análisis Horizontal y Vertical?"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   7695
      _Version        =   1572864
      _ExtentX        =   13568
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Periodos:"
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1296
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
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   312
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPeriodo 
         Height          =   312
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAnioCorte 
         Height          =   312
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   732
         _Version        =   1572864
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtMesCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   23
         Top             =   720
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPeriodoCorte 
         Height          =   312
         Left            =   2880
         TabIndex        =   24
         Top             =   720
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   550
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
      Begin VB.Label Label1 
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
         Height          =   312
         Index           =   0
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   948
      End
      Begin VB.Label Label3 
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
         Height          =   312
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   948
      End
   End
   Begin XtremeSuiteControls.CheckBox chkCuentaConMov 
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   5520
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Cuentas con Movimientos o Saldos ?"
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
      Value           =   1
      Alignment       =   1
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
      Left            =   120
      TabIndex        =   4
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
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance de Situación"
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
      Index           =   1
      Left            =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmCntX_RepBalanceSituacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBusca As Integer, vPaso As Boolean

Private Sub sbCreaReporte()
Dim strSQL As String

Me.MousePointer = vbHourglass




        
'Borra listado anterior
strSQL = "delete CNTX_REP_BALANCE_SITUACION where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and usuario = '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL, 0)


If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
    strSQL = "insert into CNTX_REP_BALANCE_SITUACION(usuario,cod_contabilidad,cod_cuenta,inicio_mes,inicio_acumulado" _
           & ",corte_mes,corte_acumulado,cod_unidad,cod_centro_costo, per_inicio_Anio, per_inicio_mes, per_Corte_anio, per_corte_mes, CLASIFICACION) " _
           & "SELECT '" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta & ",C.COD_CUENTA, isnull(SUM(X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_MES" _
           & ",isnull(SUM(X.SALDO_INICIAL + X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_ACUMULADO" _
           & ",isnull(SUM(Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_MES" _
           & ",isnull(SUM(Y.SALDO_INICIAL + Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_ACUMULADO,'',''" _
           & "," & txtAnio.Text & "," & txtMes.Text & "," & txtAnioCorte.Text & "," & txtMesCorte.Text & ", Tc.Clasificacion" _
           & " FROM CntX_Cuentas C LEFT JOIN vCntX_Mov_Cuentas_General X ON C.COD_CUENTA = X.COD_CUENTA And C.cod_contabilidad = X.cod_contabilidad" _
           & " AND X.ANIO = " & txtAnio & " AND X.MES = " & txtMes _
           & " LEFT JOIN vCntX_Mov_Cuentas_General Y ON  C.COD_CUENTA = Y.COD_CUENTA And C.cod_contabilidad = Y.cod_contabilidad" _
           & " AND Y.ANIO = " & txtAnioCorte & " AND Y.MES = " & txtMesCorte _
           & " INNER JOIN CntX_Tipos_Cuentas Tc on C.Tipo_Cuenta = Tc.Tipo_Cuenta and C.cod_Contabilidad = Tc.cod_Contabilidad" _
           & " WHERE C.Cod_Contabilidad = " & gCntX_Parametros.CodigoConta _
           & " AND C.COD_CUENTA BETWEEN '" & fxCntX_CuentaFormato(False, txtDesde) & "' AND '" & fxCntX_CuentaFormato(False, txtHasta) & "'" _
           & " group by C.COD_CUENTA, Tc.Clasificacion"
    Call ConectionExecute(strSQL, 0)

End If

If cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then
    strSQL = "insert into CNTX_REP_BALANCE_SITUACION(usuario,cod_contabilidad,cod_cuenta,inicio_mes,inicio_acumulado" _
           & ",corte_mes,corte_acumulado,cod_unidad,cod_centro_costo, per_inicio_Anio, per_inicio_mes, per_Corte_anio, per_corte_mes, CLASIFICACION) " _
           & "SELECT '" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta & ",C.COD_CUENTA, isnull(SUM(X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_MES" _
           & ",isnull(SUM(X.SALDO_INICIAL + X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_ACUMULADO" _
           & ",isnull(SUM(Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_MES" _
           & ",isnull(SUM(Y.SALDO_INICIAL + Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_ACUMULADO,'','" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
           & "," & txtAnio.Text & "," & txtMes.Text & "," & txtAnioCorte.Text & "," & txtMesCorte.Text & ", Tc.Clasificacion" _
           & " FROM CntX_Cuentas C LEFT JOIN vCntX_Mov_Cuentas_CentroCosto X ON C.COD_CUENTA = X.COD_CUENTA And C.cod_contabilidad = X.cod_contabilidad" _
           & " And X.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
           & " AND X.ANIO = " & txtAnio & " AND X.MES = " & txtMes _
           & " LEFT JOIN vCntX_Mov_Cuentas_CentroCosto Y ON  C.COD_CUENTA = Y.COD_CUENTA And C.cod_contabilidad = Y.cod_contabilidad" _
           & " and Y.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
           & " AND Y.ANIO = " & txtAnioCorte & " AND Y.MES = " & txtMesCorte _
           & " INNER JOIN CntX_Tipos_Cuentas Tc on C.Tipo_Cuenta = Tc.Tipo_Cuenta and C.cod_Contabilidad = Tc.cod_Contabilidad" _
           & " WHERE C.Cod_Contabilidad = " & gCntX_Parametros.CodigoConta _
           & " AND C.COD_CUENTA BETWEEN '" & fxCntX_CuentaFormato(False, txtDesde) & "' AND '" & fxCntX_CuentaFormato(False, txtHasta) & "'" _
           & " group by C.COD_CUENTA, Tc.Clasificacion"
    Call ConectionExecute(strSQL, 0)

End If

If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS" Then
    strSQL = "insert into CNTX_REP_BALANCE_SITUACION(usuario,cod_contabilidad,cod_cuenta,inicio_mes,inicio_acumulado" _
           & ",corte_mes,corte_acumulado,cod_unidad,cod_centro_costo, per_inicio_Anio, per_inicio_mes, per_Corte_anio, per_corte_mes, CLASIFICACION) " _
           & "SELECT '" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta & ",C.COD_CUENTA, isnull(SUM(X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_MES" _
           & ",isnull(SUM(X.SALDO_INICIAL + X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_ACUMULADO" _
           & ",isnull(SUM(Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_MES" _
           & ",isnull(SUM(Y.SALDO_INICIAL + Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_ACUMULADO,'" _
           & cboUnidad.ItemData(cboUnidad.ListIndex) & "',''" _
           & "," & txtAnio.Text & "," & txtMes.Text & "," & txtAnioCorte.Text & "," & txtMesCorte.Text & ", Tc.Clasificacion" _
           & " FROM CntX_Cuentas C LEFT JOIN vCntX_Mov_Cuentas_Unidad X ON C.COD_CUENTA = X.COD_CUENTA And C.cod_contabilidad = X.cod_contabilidad" _
           & " And X.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
           & " AND X.ANIO = " & txtAnio & " AND X.MES = " & txtMes _
           & " LEFT JOIN vCntX_Mov_Cuentas_Unidad Y ON  C.COD_CUENTA = Y.COD_CUENTA And C.cod_contabilidad = Y.cod_contabilidad" _
           & " And Y.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
           & " AND Y.ANIO = " & txtAnioCorte & " AND Y.MES = " & txtMesCorte _
           & " INNER JOIN CntX_Tipos_Cuentas Tc on C.Tipo_Cuenta = Tc.Tipo_Cuenta and C.cod_Contabilidad = Tc.cod_Contabilidad" _
           & " WHERE C.Cod_Contabilidad = " & gCntX_Parametros.CodigoConta _
           & " AND C.COD_CUENTA BETWEEN '" & fxCntX_CuentaFormato(False, txtDesde) & "' AND '" & fxCntX_CuentaFormato(False, txtHasta) & "'" _
           & " group by C.COD_CUENTA, Tc.Clasificacion"
    Call ConectionExecute(strSQL, 0)
End If

If cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS" Then

    strSQL = "insert into CNTX_REP_BALANCE_SITUACION(usuario,cod_contabilidad,cod_cuenta,inicio_mes,inicio_acumulado" _
           & ",corte_mes,corte_acumulado,cod_unidad,cod_centro_costo, per_inicio_Anio, per_inicio_mes, per_Corte_anio, per_corte_mes, CLASIFICACION) " _
           & "SELECT '" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta & ",C.COD_CUENTA, isnull(SUM(X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_MES" _
           & ",isnull(SUM(X.SALDO_INICIAL + X.TOTAL_DEBITOS + X.TOTAL_CREDITOS),0) AS X_ACUMULADO" _
           & ",isnull(SUM(Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_MES" _
           & ",isnull(SUM(Y.SALDO_INICIAL + Y.TOTAL_DEBITOS + Y.TOTAL_CREDITOS),0) AS Y_ACUMULADO,'" _
           & cboUnidad.ItemData(cboUnidad.ListIndex) & "','" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
           & "," & txtAnio.Text & "," & txtMes.Text & "," & txtAnioCorte.Text & "," & txtMesCorte.Text & ", Tc.Clasificacion" _
           & " FROM CntX_Cuentas C LEFT JOIN CntX_Mov_Cuentas_Detallado X ON C.COD_CUENTA = X.COD_CUENTA And C.cod_contabilidad = X.cod_contabilidad" _
           & " And X.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "' and X.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
           & " AND X.ANIO = " & txtAnio & " AND X.MES = " & txtMes _
           & " LEFT JOIN CntX_Mov_Cuentas_Detallado Y ON  C.COD_CUENTA = Y.COD_CUENTA And C.cod_contabilidad = Y.cod_contabilidad" _
           & " And Y.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "' and Y.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'" _
           & " AND Y.ANIO = " & txtAnioCorte & " AND Y.MES = " & txtMesCorte _
           & " INNER JOIN CntX_Tipos_Cuentas Tc on C.Tipo_Cuenta = Tc.Tipo_Cuenta and C.cod_Contabilidad = Tc.cod_Contabilidad" _
           & " WHERE C.Cod_Contabilidad = " & gCntX_Parametros.CodigoConta _
           & " AND C.COD_CUENTA BETWEEN '" & fxCntX_CuentaFormato(False, txtDesde) & "' AND '" & fxCntX_CuentaFormato(False, txtHasta) & "'" _
           & " group by C.COD_CUENTA, Tc.Clasificacion"
    Call ConectionExecute(strSQL, 0)

End If


     
Me.MousePointer = vbDefault
       
End Sub


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
Dim strSQL As String, rs As New ADODB.Recordset, strTitulo As String
Dim pInicio As String, pCorte As String


Me.MousePointer = vbHourglass

On Error GoTo vError

Dim pUnidad As String, pCentroCosto As String

If Mid(cboUnidad.Text, 1, 2) = "[C" Then
   pUnidad = "0x0"
Else
   pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
End If

If cboCentroCosto.Text = "TODOS" Then
    pCentroCosto = "0x0"
Else
    pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If

strSQL = "exec spCntX_BalanceSituacion_Procesa " & gCntX_Parametros.CodigoConta _
        & ", " & txtAnio.Text & ", " & txtMes.Text _
        & ", " & txtAnioCorte.Text & ", " & txtMesCorte.Text _
        & ",'" & fxCntX_CuentaFormato(False, txtDesde.Text) & "', '" & fxCntX_CuentaFormato(False, txtHasta) _
        & "', '" & pUnidad & "', '" & pCentroCosto & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
        
strSQL = "select dbo.fxSys_FechaAnioMesToDatetime(" & txtAnio.Text & ", " & txtMes.Text & ") as 'Inicio'" _
       & " , dbo.fxSys_FechaAnioMesToDatetime(" & txtAnioCorte.Text & ", " & txtMesCorte.Text & ") as 'Corte'"
Call OpenRecordSet(rs, strSQL)
    pInicio = Format(rs!Inicio, "yyyy-mm-dd")
    pCorte = Format(rs!Corte, "yyyy-mm-dd")
rs.Close



strSQL = "{CNTX_REP_BALANCE_SITUACION.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
       & " AND {CNTX_REP_BALANCE_SITUACION.USUARIO} = '" & glogon.Usuario & "'"
       
If chkCuentaConMov.Value = xtpUnchecked Then
strSQL = strSQL & " and ({CNTX_REP_BALANCE_SITUACION.INICIO_MES} + {CNTX_REP_BALANCE_SITUACION.INICIO_ACUMULADO} " _
       & " + {CNTX_REP_BALANCE_SITUACION.CORTE_MES} + {CNTX_REP_BALANCE_SITUACION.CORTE_ACUMULADO}) <> 0"

End If
       
       
strTitulo = " Periodos : " & txtPeriodo & " Y " & txtPeriodoCorte
strTitulo = strTitulo & " [ Cuentas : " & txtDesde & " Hasta " & txtHasta
strTitulo = strTitulo & "] [ Unidad : " & cboUnidad.Text & "] [ Centro Costo : " & cboCentroCosto.Text & "]"

If chkAnalisiHV.Value = vbUnchecked Then
    Call sbCntX_Reportes("BALANCESITUACION", strSQL, strTitulo, pInicio, pCorte, cboNiveles.Text, , pUnidad, pCentroCosto)
Else
    Call sbCntX_Reportes("BALANCESITUACION_A", strSQL, strTitulo, pInicio, pCorte, cboNiveles.Text, , pUnidad, pCentroCosto)
End If

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
vBusca = 1
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vPaso = True
Call sbCntX_CargaCboUnidades(cboUnidad)
vPaso = False


cboNiveles.AddItem "1"
cboNiveles.AddItem "2"
cboNiveles.AddItem "3"
cboNiveles.AddItem "4"
cboNiveles.AddItem "5"
cboNiveles.AddItem "6"
cboNiveles.AddItem "7"
cboNiveles.AddItem "8"

cboNiveles.Text = "4"

Call cboUnidad_Click

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



Private Sub txtAnio_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMes.SetFocus
End Sub

Private Sub txtAnioCorte_Change()
On Error GoTo vError
txtPeriodoCorte = fxCntX_PeriodoDesc(txtAnioCorte, txtMesCorte)
vError:
End Sub

Private Sub txtAnioCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMesCorte.SetFocus
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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub txtHasta_LostFocus()
 txtHastaDesc.Text = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtHasta.Text))
 txtHasta.Text = fxCntX_CuentaFormato(True, txtHasta.Text)
End Sub


Private Sub txtMes_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnioCorte.SetFocus
End Sub

Private Sub txtMesCorte_Change()
On Error GoTo vError
txtPeriodoCorte = fxCntX_PeriodoDesc(txtAnioCorte, txtMesCorte)
vError:
End Sub

Private Sub txtMesCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And cmdReporte.Enabled Then cmdReporte.SetFocus
End Sub

