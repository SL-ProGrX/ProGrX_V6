VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_RepMovTipoCuenta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analítico de Cuentas"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   HelpContextID   =   21
   Icon            =   "frmCntX_RepMovTipoCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   2400
      TabIndex        =   4
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
      Left            =   2400
      TabIndex        =   5
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1212
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   8772
      _Version        =   1572864
      _ExtentX        =   15473
      _ExtentY        =   2138
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   615
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1695
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
         Picture         =   "frmCntX_RepMovTipoCuenta.frx":000C
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   1692
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
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Informe: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   2412
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
         TabIndex        =   10
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.GroupBox fraRangos 
      Height          =   1212
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   7812
      _Version        =   1572864
      _ExtentX        =   13779
      _ExtentY        =   2138
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   17
         Top             =   720
         Width           =   612
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
         TabIndex        =   16
         Top             =   360
         Width           =   612
      End
   End
   Begin XtremeSuiteControls.ComboBox cboNiveles 
      Height          =   312
      Left            =   1680
      TabIndex        =   18
      Top             =   4680
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2355
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1680
      TabIndex        =   19
      Top             =   3960
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1680
      TabIndex        =   20
      Top             =   4320
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.CheckBox chkRango 
      Height          =   252
      Left            =   5520
      TabIndex        =   21
      Top             =   2280
      Width           =   2652
      _Version        =   1572864
      _ExtentX        =   4678
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Filtrar el rango de cuentas "
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkDivisa 
      Height          =   252
      Left            =   3600
      TabIndex        =   22
      Top             =   4680
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Movimientos en Divisa Origen"
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
   Begin XtremeSuiteControls.CheckBox chkMovimientosCero 
      Height          =   252
      Left            =   3600
      TabIndex        =   23
      Top             =   3960
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Cuentas con Cero Movimientos"
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
   Begin XtremeSuiteControls.CheckBox chkSubCuentas 
      Height          =   252
      Left            =   3600
      TabIndex        =   24
      Top             =   4320
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Todas las Sub Cuentas "
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
   Begin XtremeSuiteControls.CheckBox chkPendientes 
      Height          =   252
      Left            =   3600
      TabIndex        =   27
      Top             =   5040
      Width           =   4572
      _Version        =   1572864
      _ExtentX        =   8064
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar Movimientos de asientos No aplicados"
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
      Left            =   360
      TabIndex        =   7
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
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Analítico de Cuentas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
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
      TabIndex        =   3
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Index           =   4
      Left            =   720
      TabIndex        =   2
      Top             =   4320
      Width           =   612
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Index           =   3
      Left            =   720
      TabIndex        =   1
      Top             =   3960
      Width           =   612
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
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   9132
   End
End
Attribute VB_Name = "frmCntX_RepMovTipoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBusca As Integer
Dim vPaso As Boolean

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

Private Sub chkRango_Click()
If chkRango.Value = xtpChecked Then
  fraRangos.Enabled = True
  txtDesde.SetFocus
Else
  fraRangos.Enabled = False
End If
End Sub


Private Sub cmdReporte_Click()
Dim strSQL As String, strSQL2 As String, strSubtitulo As String
Dim rs As New ADODB.Recordset, vCuentaInicio As String, vCuentaCorte As String
'Verifica

On Error GoTo vError

If dtpInicio.Value > dtpCorte.Value Then
  MsgBox "Rango de fecha no es válido, Verifique...", vbCritical
  Exit Sub
End If




If chkRango.Value = vbChecked Then
    Select Case ""
      Case txtDesde, txtHasta
        MsgBox "Datos especificados en los rangos, no son válidos. Verifiquelos...", vbCritical
        Exit Sub
    End Select
End If


strSubtitulo = "Inicio: " _
             & Format(dtpInicio.Value, "dd/mm/yyyy") & " Corte: " & Format(dtpCorte.Value, "dd/mm/yyyy")

If chkRango.Value = vbChecked Then
   
   vCuentaInicio = fxCntX_CuentaFormato(False, txtDesde)
   vCuentaCorte = fxCntX_CuentaFormato(False, txtHasta)
   
   strSubtitulo = strSubtitulo & " ¦ Cuentas: " & txtDesde.Text _
                & " - " & txtHasta.Text
Else
  strSQL2 = "select min(cod_cuenta) as MinC, max(cod_cuenta) as MaxC from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta
  Call OpenRecordSet(rs, strSQL2)
    vCuentaInicio = Trim(rs!minc & "")
    vCuentaCorte = Trim(rs!maxc) & ""
  rs.Close

End If


If chkPendientes.Value = xtpChecked Then
  strSubtitulo = strSubtitulo & " ¦ Incluye asientos pendientes"
End If


Dim pUnidad As String, pCentroCosto As String

If Mid(cboUnidad.Text, 1, 2) = "[C" Then
   pUnidad = "0x0"
Else
   pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
   strSubtitulo = strSubtitulo & " ¦ Unidad: " & cboUnidad.Text

End If

If cboCentroCosto.Text = "TODOS" Then
    pCentroCosto = "0x0"
Else
    pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
    strSubtitulo = strSubtitulo & " ¦ C.C.: " & cboCentroCosto.Text
End If

'Procesa Saldo Inicial y Movimientos segun el Filtrado
Call sbCntX_MovimientoCuentas(dtpInicio.Value, dtpCorte.Value, vCuentaInicio, vCuentaCorte, chkMovimientosCero.Value _
                , pUnidad, pCentroCosto, chkDivisa.Value, chkPendientes.Value)

'Filtros para el Reporte
If chkSubCuentas.Value = 0 Then
    strSQL = "{CntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "' AND " _
           & "{CntX_Mov_Cuenta_Tmp.cod_contabilidad} = " & gCntX_Parametros.CodigoConta & " AND " _
           & "{CntX_Mov_Cuenta_Tmp.COD_CUENTA} >= '" & vCuentaInicio & "' AND " _
           & "{CntX_Mov_Cuenta_Tmp.COD_CUENTA} <= '" & vCuentaCorte & "'"
Else
    strSQL = "{CntX_Mov_Cuenta_Tmp.USUARIO} = '" & glogon.Usuario & "' AND " _
           & "{CntX_Mov_Cuenta_Tmp.cod_contabilidad} = " & gCntX_Parametros.CodigoConta
End If


If pUnidad = "0x0" Then
   pUnidad = ""
End If

If pCentroCosto = "0x0" Then
   pCentroCosto = ""
End If

If Mid(cboUnidad.Text, 1, 2) = "[C" Then
    'Consolidado
    'PROGRAMA EL SALDO INICIAL AL DIA ANTERIOR AL CORTE REALIZADO
    Call sbCntX_Reportes("MOVCUENTASCON", strSQL, strSubtitulo, chkDivisa.Value, cboTipo.Text, chkPendientes.Value, , pUnidad, pCentroCosto)

Else
    Call sbCntX_Reportes("MOVCUENTAS", strSQL, strSubtitulo, chkDivisa.Value, cboTipo.Text, chkPendientes.Value, , pUnidad, pCentroCosto)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

cboTipo.Clear
cboTipo.AddItem "Estandar"
cboTipo.AddItem "Referencia"
cboTipo.AddItem "Extendido"
cboTipo.Text = "Estandar"


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

vPaso = True
    Call sbCntX_CargaCboUnidades(cboUnidad)
vPaso = False

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
vBusca = 1

Call chkRango_Click
Call cboUnidad_Click

End Sub

Private Sub sbConsultas()
Select Case vBusca
  Case 1 'Desde
'     gBusquedas.Columna = "cod_cuenta"
'     gBusquedas.Orden = "cod_cuenta"
'     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
'     gBusquedas.Consulta = "select cod_cuenta,descripcion from CntX_Cuentas"
'     frmBusquedas.Show vbModal
'     txtDesde = gBusquedas.Resultado
  
     frmCntX_ConsultaCuentas.Show vbModal
     txtDesde = gCuenta
  
  Case 2 'hasta
'     gBusquedas.Columna = "cod_cuenta"
'     gBusquedas.Orden = "cod_cuenta"
'     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
'     gBusquedas.Consulta = "select cod_cuenta,descripcion from CntX_Cuentas"
'     frmBusquedas.Show vbModal
'     txtHasta = gBusquedas.Resultado

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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub txtHasta_LostFocus()
 txtHastaDesc.Text = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtHasta.Text))
 txtHasta.Text = fxCntX_CuentaFormato(True, txtHasta.Text)
End Sub


