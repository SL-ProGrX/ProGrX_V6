VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_RepMovPeriodo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informes de Movimientos del Periodo Ordinario"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ComboBox cboNiveles 
      Height          =   312
      Left            =   6240
      TabIndex        =   4
      Top             =   3000
      Width           =   2172
      _Version        =   1572864
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.ComboBox cboMostrar 
      Height          =   312
      Left            =   6240
      TabIndex        =   6
      Top             =   3360
      Width           =   2172
      _Version        =   1572864
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   3120
      TabIndex        =   8
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
      Left            =   3120
      TabIndex        =   9
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
   Begin XtremeSuiteControls.ComboBox cboArea 
      Height          =   312
      Left            =   3120
      TabIndex        =   12
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
   Begin XtremeSuiteControls.ComboBox cboReporte 
      Height          =   312
      Left            =   3120
      TabIndex        =   13
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
      Left            =   3120
      TabIndex        =   16
      Top             =   240
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
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   312
      Left            =   3120
      TabIndex        =   17
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
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   8415
      _Version        =   1572864
      _ExtentX        =   14843
      _ExtentY        =   2143
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGenerar 
         Height          =   612
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Generar"
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
         Picture         =   "frmCntX_RepMovPeriodo.frx":0000
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
         TabIndex        =   3
         Top             =   240
         Width           =   4812
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   2640
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Area de Trabajo"
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
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   2280
      Width           =   1692
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
      Left            =   1080
      TabIndex        =   11
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
      Left            =   1080
      TabIndex        =   10
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   3360
      Width           =   852
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
      Index           =   7
      Left            =   5400
      TabIndex        =   5
      Top             =   3000
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Index           =   1
      Left            =   1880
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   1880
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmCntX_RepMovPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean



Private Sub cboReporte_Click()

cboNiveles.Clear

If cboReporte.ItemData(cboReporte.ListIndex) = "03" Then
    cboNiveles.AddItem "Unidad"
    cboNiveles.AddItem "Centro de Costo"
    cboNiveles.Text = "Unidad"

Else
    cboNiveles.AddItem "1"
    cboNiveles.AddItem "2"
    cboNiveles.AddItem "3"
    cboNiveles.AddItem "4"
    cboNiveles.AddItem "5"
    cboNiveles.AddItem "6"
    cboNiveles.AddItem "7"
    cboNiveles.AddItem "8"
   
    cboNiveles.Text = "4"
End If
End Sub

Private Sub cboTipo_Click()
cboUnidad.Enabled = False
cboCentroCosto.Enabled = False
cboArea.Enabled = False

If cboTipo.ItemData(cboTipo.ListIndex) = "01" Then
    cboUnidad.Enabled = True
    cboCentroCosto.Enabled = True
Else
    cboArea.Enabled = True
End If
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


Private Sub sbMovimientoCatalogo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iMes As Integer, lngAnio As Long
Dim vMeses As Integer, vMesesNombre(12) As String
Dim vFiltro As String

'1. Borrar los datos anteriores del usuario
'2. Cargar las Cuentas a Buscar
'3. Barrer todos los meses en busca de los movimientos
'4. Preparar el Reporte

On Error GoTo vError

Me.MousePointer = vbHourglass

lbl.Caption = "Eliminando Información Anterior..."
lbl.Refresh

strSQL = "delete CntX_Rep_Periodos_mov where usuario = '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL, 0)

lbl.Caption = "Inicializando Cuentas ..."
lbl.Refresh

Select Case cboReporte.ItemData(cboReporte.ListIndex)
  Case "01" 'Ingresos
     vFiltro = "('I','V')"
  Case "02" 'Gastos
     vFiltro = "('G')"
  Case "03" 'Resultados (No Aplica
     Exit Sub
  Case "04" 'Activos
     vFiltro = "('A')"
  Case "05" 'Pasivos
     vFiltro = "('P')"
  Case "06" 'Patrimonio
     vFiltro = "('C')"
  Case "07" 'Balance Completo
     vFiltro = "('I','V','C','G','P','A')"
  Case "08" 'Activos, Pasivos y Patrimonio (General) sin Resultados
     vFiltro = "('A','P', 'C')"
  Case "09" 'Ingresos y Gastos
     vFiltro = "('I', 'V', 'G')"
End Select

If cboTipo.ItemData(cboTipo.ListIndex) = "01" Then 'Contabilidad General
  
     strSQL = "insert into CntX_Rep_Periodos_mov(cod_cuenta,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
            & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
            & ",movimiento_08,movimiento_09) (select cod_cuenta,'" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
            & ",0,0,0,0,0,0,0,0,0,0,0,0" _
            & " from CntX_Cuentas C inner join CntX_Tipos_Cuentas T on C.cod_contabilidad = T.cod_contabilidad" _
            & " and C.tipo_cuenta = T.tipo_cuenta" _
            & " where T.clasificacion in " & vFiltro & " and C.cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
  
Else 'Area de Trabajo
     strSQL = "insert into CntX_Rep_Periodos_mov(cod_cuenta,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
            & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
            & ",movimiento_08,movimiento_09) (select C.cod_cuenta,'" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
            & ",0,0,0,0,0,0,0,0,0,0,0,0" _
            & " from CntX_Cuentas C inner join CntX_Tipos_Cuentas T on C.cod_contabilidad = T.cod_contabilidad" _
            & " and C.tipo_cuenta = T.tipo_cuenta inner join CntX_Area_Cuentas A" _
            & " On C.cod_cuenta = A.cod_cuenta and C.cod_contabilidad = A.cod_contabilidad" _
            & " where T.clasificacion in " & vFiltro & " and C.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
            & " and A.cod_area = " & cboArea.ItemData(cboArea.ListIndex) & ")"
   
End If 'CboTipo

'Carga Cuentas a Utilizar
Call ConectionExecute(strSQL, 0)


strSQL = "select * from cntx_Cierres " _
       & " where ID_CIERRE = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta

Call OpenRecordSet(rs, strSQL, 0)
  iMes = rs!Inicio_Mes
  lngAnio = rs!Inicio_Anio
rs.Close


'Procesa Meses del Periodo Fiscal
For vMeses = 1 To 12
   lbl.Caption = "Procesando Periodo " & lngAnio & "-" & iMes
   lbl.Refresh
   
   vMesesNombre(vMeses) = fxCntX_MesDesc(iMes)
   
   If cboMostrar.Text = "Acumulados" Then
       strSQL = "select M.cod_cuenta,M.SALDO_Inicial + M.Total_Debitos + M.Total_Creditos as Movimiento"
   Else
       strSQL = "select M.cod_cuenta,M.Total_Debitos + M.Total_Creditos as Movimiento"
   End If
   
    
     'Selecciona la fuente de la información dependiento de los filtros
     Select Case True
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
          strSQL = strSQL & " from vCntX_Mov_Cuentas_General M inner join CntX_Rep_Periodos_mov R"
     
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          strSQL = strSQL & " from vCntX_Mov_Cuentas_CentroCosto M inner join CntX_Rep_Periodos_mov R"
            
       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
          strSQL = strSQL & " from vCntX_Mov_Cuentas_Unidad M inner join CntX_Rep_Periodos_mov R"

       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          strSQL = strSQL & " from CntX_Mov_Cuentas_Detallado M inner join CntX_Rep_Periodos_mov R"
     End Select
   
   strSQL = strSQL & " On M.cod_contabilidad = R.cod_contabilidad and M.cod_cuenta = R.cod_cuenta" _
          & " where M.anio = " & lngAnio & " and M.mes = " & iMes _
          & " and R.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
          & " and R.usuario = '" & glogon.Usuario & "'"
   
     'Filtros Finales
     Select Case True
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
     
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          strSQL = strSQL & " and M.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'"
            
       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
          strSQL = strSQL & " and M.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'"

       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          strSQL = strSQL & " and M.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                 & " and M.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'"
     End Select
   
   
   Call OpenRecordSet(rs, strSQL, 0)
   
   strSQL = ""
   Do While Not rs.EOF
    
     strSQL = strSQL & Space(10) & "update CntX_Rep_Periodos_mov set MOVIMIENTO_" _
            & Format(vMeses, "00") & " = " & rs!Movimiento _
            & " where usuario = '" & glogon.Usuario & "'" _
            & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
            & " and cod_cuenta = '" & Trim(rs!cod_cuenta) & "'"
     
     If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL, 0)
        strSQL = ""
     End If
     
     rs.MoveNext
   Loop
   rs.Close
         
    'Ejecuta el Ultimo Bloque
     If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL, 0)
        strSQL = ""
     End If


   'Periodo Siguiente
   If iMes = 12 Then
        lngAnio = lngAnio + 1
        iMes = 1
   Else
        iMes = iMes + 1
   End If

Next vMeses

lbl.Caption = "Preparando Reporte ..."
lbl.Refresh


With frmContenedor.Crt
     .Reset
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "ProGrX: Contabilidad"
     .Formulas(0) = "empresa = '" & gCntX_Parametros.NombreEmpresa & "'"
     .Formulas(1) = "fecha = 'Fecha ..: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(2) = "usuario = 'Usuario ..: " & glogon.Usuario & "'"
     
     
     If cboTipo.ItemData(cboTipo.ListIndex) = "01" Then
       .Formulas(3) = "Area = 'UNIDAD : " & cboUnidad.Text & " CENTRO DE COSTO : " & cboCentroCosto.Text & " - NIVEL ..:" _
                    & cboNiveles.Text & " - " & UCase(cboMostrar.Text) & "'"
     Else
       .Formulas(3) = "Area = '" & UCase(Mid(cboArea.Text, 6, 60)) & " - NIVEL ..: " _
                    & cboNiveles.Text & " - " & UCase(cboMostrar.Text) & "'"
     End If
     
     .Formulas(4) = "Titulo = '" & UCase(Mid(cboReporte.Text, 6, 60)) & "'"
     .Formulas(5) = "SubTitulo = '" & cboPeriodo.Text & "'"
     .Formulas(6) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
     
'     For vMeses = 1 To 12
'        .Formulas(6 + vMeses) = "fxMes" & Format(vMeses, "00") & " = '" & UCase(fxCntX_MesDesc(vMeses)) & "'"
'     Next vMeses
     
     'Titulos de los meses en el reporte
     strSQL = "select * from cntx_Cierres where id_Cierre = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
            & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     Call OpenRecordSet(rs, strSQL, 0)
          iMes = rs!Inicio_Mes
          lngAnio = rs!Inicio_Anio
     rs.Close
     For vMeses = 1 To 12
        .Formulas(6 + vMeses) = "fxMes" & Format(vMeses, "00") & " = '" & UCase(fxCntX_MesDesc(iMes)) & "'"
        If iMes = 12 Then
             iMes = 1
        Else
             iMes = iMes + 1
        End If
     Next vMeses
     
     
     .Connect = glogon.ConectRPT
     
     
     .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_MovPeriodo.rpt")
     .SelectionFormula = "{CntX_Rep_Periodos_mov.USUARIO} = '" & glogon.Usuario & "' AND {CntX_Cuentas.Nivel} <= " & cboNiveles.Text
        
    .Action = 1

End With

lbl.Caption = ""
Me.MousePointer = vbDefault


Exit Sub

vError:
    lbl.Caption = ""
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Me.MousePointer = vbDefault

End Sub


Private Sub sbMovimientoResultados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iMes As Integer, lngAnio As Long
Dim vMeses As Integer, vMesesNombre(12) As String, vMostrar As String
Dim vFiltro As String, vUnidad As String, vCentroCosto As String

'1. Borrar los datos anteriores del usuario
'2. Cargar las Cuentas a Buscar
'3. Barrer todos los meses en busca de los movimientos
'4. Preparar el Reporte

On Error GoTo vError

Me.MousePointer = vbHourglass

lbl.Caption = "Eliminando Información Anterior..."
lbl.Refresh

strSQL = "delete CNTX_REP_PERIODOS_MOV_UNIDAD where usuario = '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL, 0)


If cboUnidad.Text = "[CONSOLIDADO]" Then
   vUnidad = ""
Else
   vUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
End If

If cboCentroCosto.Text = "TODOS" Then
   vCentroCosto = ""
Else
   vCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If




If cboTipo.ItemData(cboTipo.ListIndex) = "01" Then 'Contabilidad General
     
     Select Case cboNiveles.Text
        Case "Unidad"
            Select Case True
               Case (vUnidad = "" And vCentroCosto = "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) (select cod_unidad,'','" & glogon.Usuario & "',cod_contabilidad" _
                          & ",0,0,0,0,0,0,0,0,0,0,0,0" _
                          & " from CntX_Unidades  where cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
               
               Case (vUnidad <> "" And vCentroCosto = "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) (select cod_unidad,'','" & glogon.Usuario & "',cod_contabilidad" _
                          & ",0,0,0,0,0,0,0,0,0,0,0,0" _
                          & " from CntX_Unidades  where cod_unidad = '" & vUnidad & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
                          
               Case (vUnidad <> "" And vCentroCosto <> "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) values('" & vUnidad & "','" & vCentroCosto & "','" & glogon.Usuario _
                          & "'," & gCntX_Parametros.CodigoConta & ",0,0,0,0,0,0,0,0,0,0,0,0)"
               
               Case (vUnidad = "" And vCentroCosto <> "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) (select '',cod_centro_costo,'" & glogon.Usuario & "',cod_contabilidad" _
                          & ",0,0,0,0,0,0,0,0,0,0,0,0" _
                          & " from CntX_Centro_Costos  where cod_centro_costo = '" & vCentroCosto & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
            End Select
            
      Case "Centro de Costo"
            Select Case True
               Case (vUnidad = "" And vCentroCosto = "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) (select cod_unidad,cod_centro_Costo,'" & glogon.Usuario & "',cod_contabilidad" _
                          & ",0,0,0,0,0,0,0,0,0,0,0,0" _
                          & " from CNTX_UNIDADES_CC  where cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
               
               Case (vUnidad <> "" And vCentroCosto = "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) (select cod_unidad,cod_centro_Costo,'" & glogon.Usuario & "',cod_contabilidad" _
                          & ",0,0,0,0,0,0,0,0,0,0,0,0" _
                          & " from CNTX_UNIDADES_CC  where cod_unidad = '" & vUnidad & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
                          
               Case (vUnidad <> "" And vCentroCosto <> "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) values('" & vUnidad & "','" & vCentroCosto & "','" & glogon.Usuario _
                          & "'," & gCntX_Parametros.CodigoConta & ",0,0,0,0,0,0,0,0,0,0,0,0)"
               
               Case (vUnidad = "" And vCentroCosto <> "")
                   strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
                          & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
                          & ",movimiento_08,movimiento_09) (select '',cod_centro_costo,'" & glogon.Usuario & "',cod_contabilidad" _
                          & ",0,0,0,0,0,0,0,0,0,0,0,0" _
                          & " from CntX_Centro_Costos  where cod_centro_costo = '" & vCentroCosto & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
            End Select
      
    End Select 'Nivel
     
     
Else 'Area de Trabajo
     strSQL = "insert into CNTX_REP_PERIODOS_MOV_UNIDAD(cod_unidad,cod_centro_costo,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
            & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
            & ",movimiento_08,movimiento_09) (select cod_unidad,cod_centro_costo,'" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
            & ",0,0,0,0,0,0,0,0,0,0,0,0" _
            & " from CntX_Area_Unidades where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
            & " and cod_area = " & cboArea.ItemData(cboArea.ListIndex) & ")"
   
End If 'CboTipo

'Carga Cuentas a Utilizar
Call ConectionExecute(strSQL, 0)


strSQL = "select * from cntx_Cierres where id_Cierre = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
  iMes = rs!Inicio_Mes
  lngAnio = rs!Inicio_Anio
rs.Close

If cboMostrar.Text = "Acumulados" Then
  vMostrar = "A"
Else
  vMostrar = "N"
End If


'Procesa Meses del Periodo Fiscal
For vMeses = 1 To 12
   lbl.Caption = "Procesando Periodo " & lngAnio & "-" & iMes
   lbl.Refresh
   
   vMesesNombre(vMeses) = fxCntX_MesDesc(iMes)
   
   
'fxCntX_UtilidadDetallada(@Anio int, @Mes smallint, @Contabilidad int
'                , @Unidad varchar(10) =  '', @CentroCosto varchar(10) = ''
'                , @Tipo char(1) = 'A') returns dec(18,2)
   strSQL = "select cod_unidad,cod_centro_costo,usuario,cod_contabilidad" _
          & " from CNTX_REP_PERIODOS_MOV_UNIDAD where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
   Call OpenRecordSet(rs, strSQL, 0)
   
   strSQL = ""
   Do While Not rs.EOF
     strSQL = strSQL & Space(10) & "update CNTX_REP_PERIODOS_MOV_UNIDAD set MOVIMIENTO_" & Format(vMeses, "00") _
            & " =  dbo.fxCntX_UtilidadDetallada(" & lngAnio & "," & iMes & ",cod_contabilidad,cod_unidad,cod_centro_costo,'" & vMostrar & "')" _
            & " where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & rs!COD_CONTABILIDAD _
            & " and cod_unidad = '" & Trim(rs!Cod_Unidad) & "' and cod_centro_costo = '" & Trim(rs!Cod_Centro_Costo) & "'"
     
     If Len(strSQL) > 20000 Then
         Call ConectionExecute(strSQL, 0)
         strSQL = ""
     End If
     
     rs.MoveNext
   Loop
   rs.Close
     
     'Procesa Ultimo Lote
     If Len(strSQL) > 0 Then
         Call ConectionExecute(strSQL, 0)
         strSQL = ""
     End If


   'Periodo Siguiente
   If iMes = 12 Then
        lngAnio = lngAnio + 1
        iMes = 1
   Else
        iMes = iMes + 1
   End If

Next vMeses

lbl.Caption = "Preparando Reporte ..."
lbl.Refresh


With frmContenedor.Crt
     .Reset
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "ProGrX: Contabilidad"
     .Formulas(0) = "empresa = '" & gCntX_Parametros.NombreEmpresa & "'"
     .Formulas(1) = "fecha = 'Fecha ..: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(2) = "usuario = 'Usuario ..: " & glogon.Usuario & "'"
     .Connect = glogon.ConectRPT
     
     
     If cboTipo.ItemData(cboTipo.ListIndex) = "01" Then
       .Formulas(3) = "Area = 'UNIDAD : " & cboUnidad.Text & " CENTRO DE COSTO : " & cboCentroCosto.Text & " - NIVEL ..:" _
                    & cboNiveles.Text & " - " & UCase(cboMostrar.Text) & "'"
     Else
       .Formulas(3) = "Area = '" & UCase(Mid(cboArea.Text, 6, 60)) & " - NIVEL ..: " _
                    & cboNiveles.Text & " - " & UCase(cboMostrar.Text) & "'"
     End If
     
     .Formulas(4) = "Titulo = 'RESULTADOS DEL PERIODO'"
     .Formulas(5) = "SubTitulo = '" & cboPeriodo.Text & "'"
     
         
     'Titulos de los meses en el reporte
     strSQL = "select * from cntx_Cierres where id_Cierre = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
            & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     Call OpenRecordSet(rs, strSQL, 0)
          iMes = rs!Inicio_Mes
          lngAnio = rs!Inicio_Anio
     rs.Close
     For vMeses = 1 To 12
        .Formulas(5 + vMeses) = "fxMes" & Format(vMeses, "00") & " = '" & UCase(fxCntX_MesDesc(iMes)) & "'"
        If iMes = 12 Then
             iMes = 1
        Else
             iMes = iMes + 1
        End If
     Next vMeses
     
       
     If cboNiveles.Text = "Unidad" Then
         .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_MovPeriodoUnidad.rpt")
     Else
         .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_MovPeriodoCentroCosto.rpt")
     End If
     
     .SelectionFormula = "{CNTX_REP_PERIODOS_MOV_UNIDAD.USUARIO} = '" & glogon.Usuario & "' AND {CNTX_REP_PERIODOS_MOV_UNIDAD.COD_CONTABILIDAD} = " & gCntX_Parametros.CodigoConta
        
    .Action = 1

End With

lbl.Caption = ""
Me.MousePointer = vbDefault


Exit Sub

vError:
    lbl.Caption = ""
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdGenerar_Click()



If cboReporte.ItemData(cboReporte.ListIndex) = "03" Then
  Call sbMovimientoResultados
Else
  Call sbMovimientoCatalogo
End If

End Sub

Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vPaso = True
Call sbCntX_CargaCboUnidades(cboUnidad)
vPaso = False

Call cboUnidad_Click

cboTipo.AddItem "Contabilidad General"
cboTipo.ItemData(cboTipo.ListCount - 1) = "01"
cboTipo.AddItem "Area de Trabajo"
cboTipo.ItemData(cboTipo.ListCount - 1) = "02"
cboTipo.Text = "Contabilidad General"

Call cboTipo_Click

strSQL = "select cod_area as 'IdX', rtrim(Descripcion) as 'ItmX' from CntX_Area_Definicion where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call sbCbo_Llena_New(cboArea, strSQL, False, True)
cboArea.Enabled = False


cboPeriodo.Clear
strSQL = "select id_Cierre as 'Idx', rtrim(descripcion) as 'itmX',activo" _
      & " from CntX_cierres where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  cboPeriodo.AddItem rs!itmX
  cboPeriodo.ItemData(cboPeriodo.ListCount - 1) = CStr(rs!IdX)
  
  If rs!Activo = 1 Then
    cboPeriodo.Text = rs!itmX
  End If
  rs.MoveNext
Loop
rs.Close

cboMostrar.Clear
cboMostrar.AddItem "Acumulados"
cboMostrar.AddItem "Neto del Mes"
cboMostrar.Text = "Neto del Mes"

cboReporte.Clear
cboReporte.AddItem "Ingresos del Periodo"
cboReporte.ItemData(cboReporte.ListCount - 1) = "01"
cboReporte.AddItem "Gastos del Periodo"
cboReporte.ItemData(cboReporte.ListCount - 1) = "02"
cboReporte.AddItem "Resultados?"
cboReporte.ItemData(cboReporte.ListCount - 1) = "03"
cboReporte.AddItem "Activos del Periodo"
cboReporte.ItemData(cboReporte.ListCount - 1) = "04"
cboReporte.AddItem "Pasivos del Periodo"
cboReporte.ItemData(cboReporte.ListCount - 1) = "05"
cboReporte.AddItem "Patrimonio del Periodo"
cboReporte.ItemData(cboReporte.ListCount - 1) = "06"
cboReporte.AddItem "Balance Completo"
cboReporte.ItemData(cboReporte.ListCount - 1) = "07"
cboReporte.AddItem "Activos/Pasivos y Patrimonio"
cboReporte.ItemData(cboReporte.ListCount - 1) = "08"
cboReporte.AddItem "Ingresos y Gastos"
cboReporte.ItemData(cboReporte.ListCount - 1) = "09"

cboReporte.Text = "Ingresos del Periodo"

End Sub
