VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmPres_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes Presupuestarios"
   ClientHeight    =   6564
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11688
   Icon            =   "frmPres_Reportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6564
   ScaleWidth      =   11688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkCuentasResumen 
      Height          =   372
      Left            =   7080
      TabIndex        =   23
      Top             =   5160
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Incluir Cuentas de Resumen?  "
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   840
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   264
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   2280
      Width           =   4212
      _Version        =   1245187
      _ExtentX        =   7429
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Informe del Periodo - Presupuesto"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox gbReporte 
      Height          =   972
      Left            =   0
      TabIndex        =   13
      Top             =   5520
      Width           =   12252
      _Version        =   1245187
      _ExtentX        =   21611
      _ExtentY        =   1714
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   492
         Left            =   9480
         TabIndex        =   14
         Top             =   360
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPres_Reportes.frx":030A
         ImageAlignment  =   0
      End
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   7080
      TabIndex        =   1
      Top             =   2640
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCentroCosto 
      Height          =   312
      Left            =   7080
      TabIndex        =   2
      Top             =   3000
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboModelo 
      Height          =   312
      Left            =   7080
      TabIndex        =   3
      Top             =   3840
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboContabilidad 
      Height          =   312
      Left            =   5520
      TabIndex        =   4
      Top             =   2280
      Width           =   5532
      _Version        =   1245187
      _ExtentX        =   9758
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   312
      Left            =   7080
      TabIndex        =   5
      Top             =   4200
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   264
      Index           =   1
      Left            =   480
      TabIndex        =   16
      Top             =   2640
      Width           =   4212
      _Version        =   1245187
      _ExtentX        =   7429
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Informe del Periodo - Comparativo"
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
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   264
      Index           =   2
      Left            =   480
      TabIndex        =   17
      Top             =   3120
      Width           =   4212
      _Version        =   1245187
      _ExtentX        =   7429
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Informe del Modelo - Presupuesto"
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
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   264
      Index           =   3
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   4212
      _Version        =   1245187
      _ExtentX        =   7429
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Informe del Modelo - Comparativo"
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
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   264
      Index           =   4
      Left            =   480
      TabIndex        =   19
      Top             =   4080
      Width           =   4212
      _Version        =   1245187
      _ExtentX        =   7429
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Informe de Partidas Deficitarias"
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
   End
   Begin XtremeSuiteControls.ComboBox cboAjuste 
      Height          =   312
      Left            =   7080
      TabIndex        =   20
      Top             =   4680
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   264
      Index           =   5
      Left            =   480
      TabIndex        =   22
      Top             =   4440
      Width           =   4212
      _Version        =   1245187
      _ExtentX        =   7429
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Informe de Ajustes"
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
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Ajustes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   2
      Left            =   5520
      TabIndex        =   21
      Top             =   4680
      Width           =   1692
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   372
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   5292
      _Version        =   1245187
      _ExtentX        =   9334
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Tipo de Informe"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   372
      Index           =   0
      Left            =   5280
      TabIndex        =   11
      Top             =   1440
      Width           =   6492
      _Version        =   1245187
      _ExtentX        =   11451
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   5
      Left            =   5520
      TabIndex        =   10
      Top             =   2640
      Width           =   1692
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   6
      Left            =   5520
      TabIndex        =   9
      Top             =   3000
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   0
      Left            =   5520
      TabIndex        =   8
      Top             =   3840
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   1
      Left            =   5520
      TabIndex        =   7
      Top             =   4200
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   13
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Presupuesto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7452
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmPres_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim vModelo As String, vCierreID As Long, vModeloAbierto As Boolean
Dim mCuenta As String, mUnidad As String, mCentroCosto As String

Private Sub btnInforme_Click()
MsgBox "Opciones no configuradas!", vbExclamation

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

 vPaso = True '
    strSQL = "select cod_contabilidad as 'IdX', Nombre as 'ItmX' from CNTX_Contabilidades" _
           & " order by cod_contabilidad"
    Call sbCbo_Llena_New(cboContabilidad, strSQL, False, True)
 vPaso = False
 
 Call cboContabilidad_Click
 
End Sub

Private Sub cboCentroCosto_Click()
If vPaso Then Exit Sub

End Sub

Private Sub cboContabilidad_Click()
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

vPaso = True


strSQL = "select P.cod_modelo as 'IdX' , P.DESCRIPCION as 'ItmX', Cc.Inicio_Anio" _
       & " From PRES_MODELOS P INNER JOIN PRES_MODELOS_USUARIOS Pmu on P.cod_Contabilidad = Pmu.cod_contabilidad" _
       & "  and P.cod_Modelo = Pmu.cod_Modelo and Pmu.Usuario = '" & glogon.Usuario & "'" _
       & " INNER JOIN CNTX_CIERRES Cc on P.cod_Contabilidad = Cc.cod_Contabilidad and P.ID_CIERRE = Cc.ID_CIERRE " _
       & " Where P.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & " group by P.cod_Modelo, P.Descripcion, Cc.Inicio_Anio" _
       & " order by Cc.INICIO_ANIO desc, P.Cod_Modelo"
Call sbCbo_Llena_New(cboModelo, strSQL, False, True)

vPaso = False

Call cboModelo_Click

Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboModelo_Click()
If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If cboModelo.ListCount = 0 Then Exit Sub

vModeloAbierto = False

strSQL = "select Cc.INICIO_ANIO,Cc.INICIO_MES, Cc.CORTE_ANIO, Cc.CORTE_MES, Pm.Estado" _
       & " from CNTX_CIERRES Cc inner join PRES_MODELOS Pm on Cc.COD_CONTABILIDAD = Pm.COD_CONTABILIDAD and Cc.ID_CIERRE = Pm.ID_CIERRE" _
       & " where Pm.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & " and Pm.COD_MODELO = '" & cboModelo.ItemData(cboModelo.ListIndex) & "'" _
       & " order by Cc.INICIO_ANIO desc"
Call OpenRecordSet(rs, strSQL)

If rs!Estado = "P" Then
    vModeloAbierto = True
End If

strSQL = "select dbo.fxSys_FechaAnioMesToDatetime(anio,mes) as 'ItmX'" _
       & " From dbo.fxPres_Periodo(" & rs!Inicio_Anio & "," & rs!Inicio_Mes & "," & rs!Corte_Anio & "," & rs!Corte_Mes & "," & cboContabilidad.ItemData(cboContabilidad.ListIndex) & ")"
rs.Close

'Call sbCbo_Llena_New(cboPeriodo, strSQL, False, False)

cboPeriodo.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboPeriodo.AddItem rs!itmX & ""

 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboPeriodo.Text = rs!itmX & ""
End If
rs.Close

vPaso = True

strSQL = "exec spPres_Modelo_Unidades " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & glogon.Usuario & "'"

Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)

If cboUnidad.ListCount > 0 Then
    cboUnidad.AddItem "CONSOLIDADO"
    cboUnidad.ItemData(cboUnidad.ListCount - 1) = "CONSOLIDADO"
    
    cboUnidad.Text = "CONSOLIDADO"
End If

strSQL = "exec spPres_Modelo_Ajustes_Permitidos " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & glogon.Usuario & "'"

Call sbCbo_Llena_New(cboAjuste, strSQL, True, True)

vPaso = False
Call cboUnidad_Click

Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboUnidad_Click()
If vPaso Then Exit Sub


Dim strSQL As String, pUnidad As String

On Error GoTo vError

vPaso = True


If cboUnidad.Text = "CONSOLIDADO" Then
   pUnidad = "CONS"
Else
   pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
End If

strSQL = "EXEc spPres_Modelo_Unidades_CC " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & pUnidad & "'"
Call sbCbo_Llena_New(cboCentroCosto, strSQL, True, True)


If cboCentroCosto.ListCount > 0 Then
    cboCentroCosto.AddItem "CONSOLIDADO"
    cboCentroCosto.ItemData(cboCentroCosto.ListCount - 1) = "CONSOLIDADO"
    
    cboCentroCosto.Text = "CONSOLIDADO"
End If


vPaso = False



Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 12

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


End Sub

