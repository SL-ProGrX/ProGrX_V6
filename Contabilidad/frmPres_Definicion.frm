VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPres_Definicion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición Inicial del Presupuesto por Partida"
   ClientHeight    =   6828
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   11292
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6828
   ScaleWidth      =   11292
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4932
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   11292
      _Version        =   1245187
      _ExtentX        =   19918
      _ExtentY        =   8700
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Presupuesto"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "txtValor"
      Item(0).Control(1)=   "cboProyectar"
      Item(0).Control(2)=   "Label4(1)"
      Item(0).Control(3)=   "Label4(2)"
      Item(0).Control(4)=   "btnProyectar"
      Item(0).Control(5)=   "btnAplicar"
      Item(0).Control(6)=   "txtPresupuesto"
      Item(0).Control(7)=   "vGrid"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4452
         Left            =   -69880
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1245187
         _ExtentX        =   19494
         _ExtentY        =   7853
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3372
         Left            =   4320
         TabIndex        =   14
         Top             =   840
         Width           =   6732
         _Version        =   524288
         _ExtentX        =   11875
         _ExtentY        =   5948
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   3
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmPres_Definicion.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboProyectar 
         Height          =   312
         Left            =   1800
         TabIndex        =   15
         Top             =   840
         Width           =   2172
         _Version        =   1245187
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtValor 
         Height          =   312
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   2172
         _Version        =   1245187
         _ExtentX        =   3831
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnProyectar 
         Height          =   372
         Left            =   1800
         TabIndex        =   19
         Top             =   1680
         Width           =   2172
         _Version        =   1245187
         _ExtentX        =   3831
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Realizar Proyección"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPres_Definicion.frx":05D6
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Cancel          =   -1  'True
         Height          =   372
         Left            =   4320
         TabIndex        =   20
         Top             =   4320
         Width           =   2172
         _Version        =   1245187
         _ExtentX        =   3831
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplicar Cambios"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPres_Definicion.frx":0CEF
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtPresupuesto 
         Height          =   312
         Left            =   8160
         TabIndex        =   21
         Top             =   4320
         Width           =   2652
         _Version        =   1245187
         _ExtentX        =   4678
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   972
         _Version        =   1245187
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   972
         _Version        =   1245187
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Proyección"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10680
      TabIndex        =   0
      Top             =   1356
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   3732
      _Version        =   1245187
      _ExtentX        =   6583
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCentroCosto 
      Height          =   312
      Left            =   7440
      TabIndex        =   2
      Top             =   600
      Width           =   3732
      _Version        =   1245187
      _ExtentX        =   6583
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboModelo 
      Height          =   312
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   3732
      _Version        =   1245187
      _ExtentX        =   6583
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboContabilidad 
      Height          =   312
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   3732
      _Version        =   1245187
      _ExtentX        =   6583
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   1320
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3720
      TabIndex        =   12
      Top             =   1320
      Width           =   6852
      _Version        =   1245187
      _ExtentX        =   12086
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cuenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Negocio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Costo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   6
      Left            =   5400
      TabIndex        =   7
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   13
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPres_Definicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean, vPeriodo As Boolean
Dim vModelo As String, vCierreID As Long, vModeloAbierto As Boolean



'Private Function fxPeriodoEstado(vAnio As Integer, vMes As Integer) As String
'Dim strSQL As String, rs As New ADODB.Recordset
'
'On Error GoTo vError
'
'strSQL = "Select estado from periodos where anio = " & vAnio _
'       & " and mes = " & vMes & " and COD_CONTABILIDAD = " & vCodEmpresa
'Call OpenRecordSet(rs, strSQL, 0)
'If rs.EOF And rs.BOF Then
' fxPeriodoEstado = "N"
'Else
' fxPeriodoEstado = rs!estado
'End If
'rs.Close
'
'Exit Function
'
'vError:
'  fxPeriodoEstado = "N"
'
'End Function
'


'Private Sub cmdAplicar_Click()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim i As Integer, vAnio As Integer, vMes As Integer
'Dim vCuenta As String
'
'On Error GoTo vError
'
'Me.MousePointer = vbHourglass
'
'vCuenta = fxFormatoCuenta(False, txtCuenta, 0)
'
'With vGrid
' For i = 1 To .MaxRows
'   .Row = i
'   .Col = 1
'   vAnio = .Value
'   .Col = 2
'   vMes = .Value
'   .Col = 3
'
'   'Solo Puede Definir el Presupuesto no Modificarlo
'   'Si se desea modificar tiene que realizarse con Asientos Presupuestarios
'
'   strSQL = "select coalesce(count(*),0) as Existe from presupuesto" _
'          & " where cod_cuenta = '" & vCuenta & "' and Anio = " & vAnio _
'          & " and Mes = " & vMes & " and COD_CONTABILIDAD = " & vCodEmpresa
'   Call OpenRecordSet(rs, strSQL, 0)
'   If rs!Existe = 0 Then
'      strSQL = "insert presupuesto(COD_CONTABILIDAD,cod_cuenta,presu_original" _
'             & ",ajuste_positivo,ajuste_negativo,presu_actual,anio,mes) values(" _
'             & vCodEmpresa & ",'" & vCuenta & "'," & CCur(.Text) & ",0,0," _
'             & CCur(.Text) & "," & vAnio & "," & vMes & ")"
'      If CCur(.Text) > 0 Then Call ConectionExecute(strSQL, 0)
'   End If
'   rs.Close
'
' Next i
'End With
'
'Call cbo_Click
'Me.MousePointer = vbDefault
'MsgBox "Presupuesto Definido Satisfactoriamente...", vbInformation
'
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'End Sub

Private Sub cboContabilidad_Click()
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

vPaso = True

strSQL = "select P.cod_modelo as 'IdX' , P.DESCRIPCION as 'ItmX'" _
       & " From PRES_MODELOS P INNER JOIN PRES_MODELOS_USUARIOS Pmu on P.cod_Contabilidad = Pmu.cod_contabilidad" _
       & "  and P.cod_Modelo = Pmu.cod_Modelo and Pmu.Usuario = '" & glogon.Usuario & "'" _
       & " Where P.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)
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

vModeloAbierto = False

strSQL = "select Cc.INICIO_ANIO,Cc.INICIO_MES, Cc.CORTE_ANIO, Cc.CORTE_MES, Pm.Estado" _
       & " from CNTX_CIERRES Cc inner join PRES_MODELOS Pm on Cc.COD_CONTABILIDAD = Pm.COD_CONTABILIDAD and Cc.ID_CIERRE = Pm.ID_CIERRE" _
       & " where Pm.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & " and Pm.COD_MODELO = '" & cboModelo.ItemData(cboModelo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Estado = "P" Then
    vModeloAbierto = True
End If


vPaso = True

strSQL = "exec spPres_Modelo_Unidades " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & glogon.Usuario & "'"

Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)


vPaso = False
Call cboUnidad_Click

vGrid.MaxRows = 0

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
    cboCentroCosto.Text = "CONSOLIDADO"
End If


vPaso = False


Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
 
 strSQL = "select Top 1 Cod_Cuenta_Mask,Descripcion from CntX_cuentas" _
        & " Where COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
        & " and Acepta_Movimientos = 1"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and Cod_Cuenta_Mask > '" & txtCuenta.Text & "' order by Cod_Cuenta_Mask asc"
    Else
       strSQL = strSQL & " and Cod_Cuenta_Mask < '" & txtCuenta.Text & "' order by Cod_Cuenta_Mask desc"
    End If

    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtCuenta.Text = rs!Cod_Cuenta_Mask
      txtDescripcion.Text = rs!Descripcion
      Call sbPresupuesto
    End If
    rs.Close
End If

vScroll = False
    FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub btnProyectar_Click()
Dim i As Integer, curMonto As Currency
Dim curValor As Currency

On Error GoTo vError

curValor = txtValor.Text

For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.Col = 2
 curMonto = CCur(vGrid.Text)

    Select Case cboProyectar.ItemData(cboProyectar.ListIndex)
      Case "Man" 'Mantener
         'Nada
      Case "Por" 'Adiciona Porcentaje
        curMonto = curMonto * ((curValor / 100) + 1)
      Case "Mnt" 'Adiciona Monto
        curMonto = curMonto + curValor
    End Select

 vGrid.Col = 3
 vGrid.Text = Format(curMonto, "Standard")

Next i

Call sbTotales


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Cuenta"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cod_Cuenta_Mask"
  gBusquedas.Orden = "Cod_Cuenta_Mask"
  gBusquedas.Consulta = "select Cod_Cuenta_Mask,Descripcion from CntX_cuentas"
  gBusquedas.Filtro = " and Acepta_Movimientos = 1" _
                    & " and COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)
  
  frmBusquedas.Show vbModal
  txtCuenta.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
  
  Call sbPresupuesto
  
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuenta.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Cuenta"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cod_Cuenta_Mask"
  gBusquedas.Orden = "Cod_Cuenta_Mask"
  gBusquedas.Consulta = "select Cod_Cuenta_Mask,Descripcion from CntX_cuentas"
  gBusquedas.Filtro = " and Acepta_Movimientos = 1" _
                    & " and COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)
  
  frmBusquedas.Show vbModal
  txtCuenta.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
  
  Call sbPresupuesto
  
End If

End Sub

Private Sub sbTotales()
Dim curMonto As Currency, i As Integer

On Error GoTo vError


curMonto = 0

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 3
  curMonto = curMonto + vGrid.Value
Next i

txtPresupuesto.Text = Format(curMonto, "Standard")


Exit Sub

vError:

End Sub

Private Sub Form_Load()
vModulo = 12

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True


vPaso = False

cboProyectar.AddItem "Mantener"
cboProyectar.ItemData(cboProyectar.ListCount - 1) = "Man"
cboProyectar.AddItem "Porcentual"
cboProyectar.ItemData(cboProyectar.ListCount - 1) = "Por"
cboProyectar.AddItem "Monto"
cboProyectar.ItemData(cboProyectar.ListCount - 1) = "Mnt"

cboProyectar.Text = "Mantener"

tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Presupuesto
    
    Case 1 'Historico
        Call sbHistorico
End Select
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

 vPaso = True
    strSQL = "select cod_contabilidad as 'IdX', Nombre as 'ItmX' from CNTX_Contabilidades"
    Call sbCbo_Llena_New(cboContabilidad, strSQL, False, True)
 vPaso = False
 
 Call cboContabilidad_Click
End Sub



Private Sub sbPresupuesto()
Dim strSQL As String, rs As New ADODB.Recordset

Dim i As Long, x As Integer

Dim pCuenta As String
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String

Dim iMes As Integer, iAnio As Integer, pPeriodo As Date

Dim vVista As String, itmX As ListViewItem


On Error GoTo vError

Me.MousePointer = vbHourglass

pCuenta = fxCntX_CuentaFormato(False, txtCuenta.Text, 0)

pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)
pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

If cboCentroCosto.ListCount = 0 Then
  pCentroCosto = ""
Else
  pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If


If cboUnidad.Text = "CONSOLIDADO" And cboCentroCosto.Text = "CONSOLIDADO" Then
  vVista = "G"
Else
  vVista = "U"
  If cboCentroCosto.Text <> "CONSOLIDADO" Then
    vVista = "C"
  End If
End If 'Consolidado



strSQL = "exec spPres_VistaPresupuesto_Cuenta " & pContabilidad & ",'" _
       & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta & "','" & vVista & "'"
       
       
Dim curAcumulado As Currency, curReal_Acumulado As Currency, curDiferencia_Acumulada As Currency, curEjecutado_Acumulado As Currency
Dim curAjustePositivo As Currency, curAjusteNegativo As Currency, curPreMensualInicial As Currency
Dim curDiferenciaTotal As Currency, curEjecutadoTotal As Currency


curAcumulado = 0
curReal_Acumulado = 0
curDiferencia_Acumulada = 0
curEjecutado_Acumulado = 0
curAjustePositivo = 0
curAjusteNegativo = 0
curDiferenciaTotal = 0
curEjecutadoTotal = 0
curPreMensualInicial = 0

With vGrid
    .MaxRows = 0
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1
        .Text = Format(rs!Periodo, "yyyy-mm-dd")
        .Col = 2
        .Text = Format(rs!REAL_ACUMULADO, "Standard")
        .Col = 3
        .Text = Format(rs!acumulado, "Standard")
     
     rs.MoveNext
    Loop
    rs.Close

End With

Call sbTotales

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbHistorico()
Dim strSQL As String, rs As New ADODB.Recordset

Dim i As Long, x As Integer

Dim pCuenta As String
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String

Dim iMes As Integer, iAnio As Integer, pPeriodo As Date

Dim vVista As String, itmX As ListViewItem


On Error GoTo vError

Me.MousePointer = vbHourglass

pCuenta = fxCntX_CuentaFormato(False, txtCuenta.Text, 0)

pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)
pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

If cboCentroCosto.ListCount = 0 Then
  pCentroCosto = ""
Else
  pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If


If cboUnidad.Text = "CONSOLIDADO" And cboCentroCosto.Text = "CONSOLIDADO" Then
  vVista = "G"
Else
  vVista = "U"
  If cboCentroCosto.Text <> "CONSOLIDADO" Then
    vVista = "C"
  End If
End If 'Consolidado



'Presupuesto del Periodo
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Periodo", 1240
lsw.ColumnHeaders.Add , , "Presupuesto", 2440, vbRightJustify
lsw.ColumnHeaders.Add , , "Real", 2440, vbRightJustify
lsw.ColumnHeaders.Add , , "Diferencia", 2440, vbRightJustify
lsw.ColumnHeaders.Add , , "( % )", 900, vbCenter
lsw.ColumnHeaders.Add , , "( + ) Ajuste", 2240, vbRightJustify
lsw.ColumnHeaders.Add , , "( - ) Ajuste", 2240, vbRightJustify
lsw.ColumnHeaders.Add , , "Original", 2440, vbRightJustify
lsw.ColumnHeaders.Add , , "Dif.Total", 2440, vbRightJustify
lsw.ColumnHeaders.Add , , "( % Total )", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Unidad", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Centro", 1000, vbCenter


strSQL = "exec spPres_VistaPresupuesto_Cuenta " & pContabilidad & ",'" _
       & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta & "','" & vVista & "'"
       
       
Dim curAcumulado As Currency, curReal_Acumulado As Currency, curDiferencia_Acumulada As Currency, curEjecutado_Acumulado As Currency
Dim curAjustePositivo As Currency, curAjusteNegativo As Currency, curPreMensualInicial As Currency
Dim curDiferenciaTotal As Currency, curEjecutadoTotal As Currency


curAcumulado = 0
curReal_Acumulado = 0
curDiferencia_Acumulada = 0
curEjecutado_Acumulado = 0
curAjustePositivo = 0
curAjusteNegativo = 0
curDiferenciaTotal = 0
curEjecutadoTotal = 0
curPreMensualInicial = 0

With lsw
    .ListItems.Clear
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , Format(rs!Periodo, "yyyy-mm-dd"))
          itmX.SubItems(1) = Format(rs!acumulado, "Standard")
          itmX.SubItems(2) = Format(rs!REAL_ACUMULADO, "Standard")
          itmX.SubItems(3) = Format(rs!DIFERENCIA_ACUMULADA, "Standard")
          itmX.SubItems(4) = Format(rs!EJECUTADO_ACUMULADO * 100, "Standard")
          itmX.SubItems(5) = Format(rs!AJUSTE_POSITIVO, "Standard")
          itmX.SubItems(6) = Format(rs!AJUSTE_NEGATIVO, "Standard")
          itmX.SubItems(7) = Format(rs!PRE_MENSUAL_INICIAL, "Standard")
          itmX.SubItems(8) = Format(rs!DIFERENCIA_TOTAL, "Standard")
          itmX.SubItems(9) = Format(rs!EJECUTADO_TOTAL * 100, "Standard")
     
          itmX.SubItems(10) = rs!cod_unidad & ""
          itmX.SubItems(11) = rs!cod_centro_costo & ""
     
     
            curAcumulado = rs!acumulado
            curReal_Acumulado = rs!REAL_ACUMULADO
            curDiferencia_Acumulada = rs!DIFERENCIA_ACUMULADA
            curEjecutado_Acumulado = 0 'EJECUTADO_ACUMULADO * 100
            curAjustePositivo = curAjustePositivo + rs!AJUSTE_POSITIVO
            curAjusteNegativo = curAjusteNegativo + rs!AJUSTE_NEGATIVO
            curDiferenciaTotal = curDiferenciaTotal + rs!DIFERENCIA_TOTAL
            curEjecutadoTotal = 0 'rs!EJECUTADO_TOTAL * 100
            curPreMensualInicial = curPreMensualInicial + rs!PRE_MENSUAL_INICIAL
     rs.MoveNext
    Loop
    rs.Close

      Set itmX = .ListItems.Add(, , "TOTALES:")
          itmX.SubItems(1) = Format(curAcumulado, "Standard")
          itmX.SubItems(2) = Format(curReal_Acumulado, "Standard")
          itmX.SubItems(3) = Format(curDiferencia_Acumulada, "Standard")
          itmX.SubItems(4) = Format(curEjecutado_Acumulado, "Standard")
          itmX.SubItems(5) = Format(curAjustePositivo, "Standard")
          itmX.SubItems(6) = Format(curAjusteNegativo, "Standard")
          itmX.SubItems(7) = Format(curPreMensualInicial, "Standard")
          itmX.SubItems(8) = Format(curDiferenciaTotal, "Standard")
          itmX.SubItems(9) = Format(curEjecutadoTotal, "Standard")

End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


