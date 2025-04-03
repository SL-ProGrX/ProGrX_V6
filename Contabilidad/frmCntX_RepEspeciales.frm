VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCntX_RepEspeciales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informes Especiales (Análisis Trimestral)"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1212
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   8892
      _Version        =   1441793
      _ExtentX        =   15684
      _ExtentY        =   2138
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGenerar 
         Height          =   612
         Left            =   7080
         TabIndex        =   3
         Top             =   360
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Generar"
         BackColor       =   -2147483633
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
         Picture         =   "frmCntX_RepEspeciales.frx":0000
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
         TabIndex        =   4
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.ComboBox cboNiveles 
      Height          =   312
      Left            =   6000
      TabIndex        =   5
      Top             =   3000
      Width           =   2172
      _Version        =   1441793
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
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   5292
      _Version        =   1441793
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
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   5292
      _Version        =   1441793
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
      Left            =   2880
      TabIndex        =   9
      Top             =   720
      Width           =   5292
      _Version        =   1441793
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
      Left            =   2880
      TabIndex        =   12
      Top             =   2520
      Width           =   5292
      _Version        =   1441793
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
      Left            =   840
      TabIndex        =   13
      Top             =   2520
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
      Left            =   840
      TabIndex        =   11
      Top             =   1560
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
      Left            =   840
      TabIndex        =   10
      Top             =   1920
      Width           =   1692
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
      Left            =   5040
      TabIndex        =   6
      Top             =   3000
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes Especiales"
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
      Height          =   495
      Index           =   0
      Left            =   1880
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmCntX_RepEspeciales"
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
    cboNiveles.Text = "2"
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
Dim vCtaUtilidad As String
Dim vUnidad As String, vCentroCosto As String

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


strSQL = "select Cuenta_GanPer From CNTX_CIERRES" _
       & " where ID_CIERRE = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
 vCtaUtilidad = Trim(rs!cuenta_ganper)
rs.Close

'cboReporte.AddItem "1   - Informe de Activos y Pasivos"
'cboReporte.AddItem "2   - Informe de Rentabilidad"
'cboReporte.AddItem "3   - Balance de Comprobación"

Select Case cboReporte.ItemData(cboReporte.ListIndex)
  Case "1", "1.1" 'Activos y Pasivos
     vFiltro = "('A','P','C')"
    'Cuenta para Utilidad Temporal
    strSQL = "insert into CntX_Rep_Periodos_mov(cod_cuenta,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
           & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
           & ",movimiento_08,movimiento_09) values('" & vCtaUtilidad & "','" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
           & ",0,0,0,0,0,0,0,0,0,0,0,0)"
    Call ConectionExecute(strSQL, 0)
  
  Case "2"
     vFiltro = "('I','V','G')"
    'Cuenta para Utilidad Temporal
    strSQL = "insert into CntX_Rep_Periodos_mov(cod_cuenta,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
           & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
           & ",movimiento_08,movimiento_09) values('" & vCtaUtilidad & "','" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
           & ",0,0,0,0,0,0,0,0,0,0,0,0)"
    Call ConectionExecute(strSQL, 0)
  
  Case "3" 'Balance Completo
     vFiltro = "('I','V','C','G','P','A')"

    'Cuenta para Utilidad Temporal
    strSQL = "insert into CntX_Rep_Periodos_mov(cod_cuenta,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
           & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
           & ",movimiento_08,movimiento_09) values('" & vCtaUtilidad & "','" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
           & ",0,0,0,0,0,0,0,0,0,0,0,0)"
    Call ConectionExecute(strSQL, 0)


End Select

    'Carga Cuentas a Utilizar
    strSQL = "insert into CntX_Rep_Periodos_mov(cod_cuenta,usuario,cod_contabilidad,movimiento_10,movimiento_11,movimiento_12" _
           & ",movimiento_01,movimiento_02,movimiento_03,movimiento_04,movimiento_05,movimiento_06,movimiento_07" _
           & ",movimiento_08,movimiento_09) (select cod_cuenta,'" & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta _
           & ",0,0,0,0,0,0,0,0,0,0,0,0" _
           & " from CntX_Cuentas C inner join CntX_Tipos_Cuentas T on C.cod_contabilidad = T.cod_contabilidad" _
           & " and C.tipo_cuenta = T.tipo_cuenta" _
           & " where T.clasificacion in " & vFiltro & " and C.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and C.cod_cuenta <> '" & vCtaUtilidad & "')"
    Call ConectionExecute(strSQL, 0)
 



'Inicializacion del Periodo
lngAnio = gCntX_Parametros.PeriodoAnio
iMes = gCntX_Parametros.PeriodoMes

For vMeses = 1 To 2
    If iMes = 1 Then
         lngAnio = lngAnio - 1
         iMes = 12
    Else
         iMes = iMes - 1
    End If
Next vMeses

'Procesa ultimo Trimestre
For vMeses = 1 To 3
   lbl.Caption = "Procesando Periodo " & lngAnio & "-" & iMes
   lbl.Refresh
   
   vMesesNombre(vMeses) = fxCntX_MesDesc(iMes)
   
   strSQL = "select M.cod_cuenta, sum(M.SALDO_Inicial + M.Total_Debitos + M.Total_Creditos) as 'Acumulado'" _
          & ", sum(M.Total_Debitos + M.Total_Creditos) as 'Movimiento'"

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
   
   strSQL = strSQL & " group by M.COD_CUENTA"
   
   Call OpenRecordSet(rs, strSQL, 0)
   
   strSQL = ""
   
   Do While Not rs.EOF
     strSQL = strSQL & Space(10) & "update CntX_Rep_Periodos_mov set MOVIMIENTO_" _
            & Format(vMeses, "00") & " = " & rs!Movimiento & ",MOVIMIENTO_" _
            & Format(vMeses + 3, "00") & " = " & rs!acumulado _
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
   
   If Len(strSQL) > 0 Then
      Call ConectionExecute(strSQL, 0)
   End If



   'Registro de la Utilidad/Excedentes del Periodo
     strSQL = "update CntX_Rep_Periodos_mov set MOVIMIENTO_" & Format(vMeses, "00") & " =  MOVIMIENTO_" & Format(vMeses, "00") _
            & " + dbo.fxCntX_UtilidadMes(" & lngAnio & "," & iMes & "," & gCntX_Parametros.CodigoConta & ",'" & vUnidad & "','" & vCentroCosto & "')" _
            & ",MOVIMIENTO_" & Format(vMeses + 3, "00") & " = MOVIMIENTO_" & Format(vMeses + 3, "00") _
            & " + dbo.fxCntX_Utilidad(" & lngAnio & "," & iMes & "," & gCntX_Parametros.CodigoConta & ",'" & vUnidad & "','" & vCentroCosto & "')" _
            & " where usuario = '" & glogon.Usuario & "'" _
            & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
            & " and cod_cuenta in(select cuenta  from dbo.fxCntX_CuentasCascada(" & gCntX_Parametros.CodigoConta & ",'" & vCtaUtilidad & "'))"
     Call ConectionExecute(strSQL, 0)


   'Periodo Siguiente
   If iMes = 12 Then
        lngAnio = lngAnio + 1
        iMes = 1
   Else
        iMes = iMes + 1
   End If

Next vMeses


strSQL = "delete CntX_Rep_Periodos_mov where usuario = '" & glogon.Usuario & "'" _
       & " and Movimiento_01 + Movimiento_02 + Movimiento_03 + Movimiento_04 + Movimiento_05 + Movimiento_06 = 0 "
Call ConectionExecute(strSQL, 0)

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
     
     
    .Formulas(3) = "Area = 'Unidad: " & cboUnidad.Text & ", Centro de Costo: " & cboCentroCosto.Text & ", Nivel:" _
                 & cboNiveles.Text & "'"

     .Formulas(4) = "Titulo = '" & cboReporte.Text & "'"
     .Formulas(5) = "SubTitulo = '" & cboPeriodo.Text & "'"
     .Formulas(6) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
     
     
    
    'Inicializacion del Periodo
    lngAnio = gCntX_Parametros.PeriodoAnio
    iMes = gCntX_Parametros.PeriodoMes
    For vMeses = 1 To 2
        If iMes = 1 Then
             lngAnio = lngAnio - 1
             iMes = 12
        Else
             iMes = iMes - 1
        End If
    Next vMeses
     
    'Encabezados
     For vMeses = 1 To 3
        .Formulas(6 + vMeses) = "fxMes" & Format(vMeses, "00") & " = '" & UCase(fxCntX_MesDesc(iMes)) & "'"
        .Formulas(6 + vMeses + 3) = "fxMes" & Format(vMeses + 3, "00") & " = '" & UCase(fxCntX_MesDesc(iMes)) & "'"
        
           'Periodo Siguiente
            If iMes = 12 Then
                 lngAnio = lngAnio + 1
                 iMes = 1
            Else
                 iMes = iMes + 1
            End If
     Next vMeses
     
     
     .Connect = glogon.ConectRPT
     
     
     
    '1   - Informe de Activos y Pasivos"
    '2   - Informe de Rentabilidad"
    '3   - Balance de Comprobación"
    
    
    
    Select Case cboReporte.ItemData(cboReporte.ListIndex)
      Case "1" 'Activos y Pasivos
         .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EspecialActivosPasivos.rpt")
      Case "2"   'Rentabilidad
         If cboCentroCosto.Text = "TODOS" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EspecialRentabilidad.rpt")
         Else
            .Formulas(7) = "CentroCosto='" & cboCentroCosto.Text & "'"
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EspecialRentabilidadCc.rpt")
         End If
      Case "3" 'Balance Completo
         .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EspecialBalance.rpt")

    End Select
     
     
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


Private Sub sbRentabilidadEspecial()
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


    
Select Case cboReporte.ItemData(cboReporte.ListIndex)
    Case "2.2" 'Por Unidad
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

  Case "2.1" 'Centro de Costo
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
  
End Select 'Reporte
     


'Carga Cuentas a Utilizar
Call ConectionExecute(strSQL, 0)

'Retrocede Periodos para Iniciar tres meses atrás
iMes = gCntX_Parametros.PeriodoMes
lngAnio = gCntX_Parametros.PeriodoAnio
For vMeses = 1 To 2
   If iMes = 1 Then
        lngAnio = lngAnio - 1
        iMes = 12
   Else
        iMes = iMes - 1
   End If
Next vMeses

'Procesa Ultimo Trimestre
For vMeses = 1 To 3
   lbl.Caption = "Procesando Periodo " & lngAnio & "-" & iMes
   lbl.Refresh
   
   strSQL = "select cod_unidad,cod_centro_costo,usuario,cod_contabilidad" _
          & " from CNTX_REP_PERIODOS_MOV_UNIDAD where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
   Call OpenRecordSet(rs, strSQL, 0)
   
   strSQL = ""
   
   Do While Not rs.EOF
     strSQL = strSQL & Space(10) & "update CNTX_REP_PERIODOS_MOV_UNIDAD set MOVIMIENTO_" & Format(vMeses, "00") _
            & " =  dbo.fxCntX_UtilidadDetallada(" & lngAnio & "," & iMes & ",cod_contabilidad,cod_unidad,cod_centro_costo,'N')" _
            & " where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & rs!COD_CONTABILIDAD _
            & " and cod_unidad = '" & Trim(rs!Cod_Unidad) & "' and cod_centro_costo = '" & Trim(rs!Cod_Centro_Costo) & "'"
     
     If Len(strSQL) > 20000 Then
         Call ConectionExecute(strSQL, 0)
         strSQL = ""
     End If
     rs.MoveNext
   Loop
   rs.Close

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

'Acumulado al Cierre
strSQL = "select cod_unidad,cod_centro_costo,usuario,cod_contabilidad" _
       & " from CNTX_REP_PERIODOS_MOV_UNIDAD where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)

strSQL = ""

Do While Not rs.EOF
   strSQL = strSQL & Space(10) & "update CNTX_REP_PERIODOS_MOV_UNIDAD set MOVIMIENTO_04" _
         & " =  dbo.fxCntX_UtilidadDetallada(" & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes & ",cod_contabilidad,cod_unidad,cod_centro_costo,'A')" _
         & " where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & rs!COD_CONTABILIDAD _
         & " and cod_unidad = '" & Trim(rs!Cod_Unidad) & "' and cod_centro_costo = '" & Trim(rs!Cod_Centro_Costo) & "'"
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL, 0)
        strSQL = ""
    End If
  rs.MoveNext
Loop
rs.Close

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL, 0)
   strSQL = ""
End If



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
     
     
     .Formulas(3) = "Area = 'UNIDAD : " & cboUnidad.Text & " CENTRO DE COSTO : " & cboCentroCosto.Text & " - NIVEL ..:" _
                  & cboNiveles.Text & "'"

     
     .Formulas(4) = "Titulo = 'RESULTADOS DEL PERIODO'"
     .Formulas(5) = "SubTitulo = '" & cboPeriodo.Text & "'"
     
         
     'Titulos de los meses en el reporte
     'Retrocede Periodos para Iniciar tres meses atrás
     iMes = gCntX_Parametros.PeriodoMes
     lngAnio = gCntX_Parametros.PeriodoAnio
     For vMeses = 1 To 2
        If iMes = 1 Then
             lngAnio = lngAnio - 1
             iMes = 12
        Else
             iMes = iMes - 1
        End If
     Next vMeses

     For vMeses = 1 To 3
        .Formulas(5 + vMeses) = "fxMes" & Format(vMeses, "00") & " = '" & UCase(fxCntX_MesDesc(iMes)) & "'"
        If iMes = 12 Then
             iMes = 1
        Else
             iMes = iMes + 1
        End If
     Next vMeses
     .Formulas(5 + 4) = "fxMes" & Format(4, "00") & " = 'ACUMULADO'"
     
     Select Case cboReporte.ItemData(cboReporte.ListIndex)
        Case "2.1" 'Rentabilidad por Centro de Costo
             .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EspecialRentabilidadCentroCosto.rpt")
        Case "2.2" 'Rentabilidad por Unidad
             .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EspecialRentabilidadUnidad.rpt")
        
     End Select
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

    '1   - Informe de Activos y Pasivos"
    '2   - Informe de Rentabilidad"
    '2.1 - Rentabilidad (Centro de Costo)"
    '2.2 - Rentabilidad (Unidad)"
    '3   - Balance de Comprobación"
    
    Select Case cboReporte.ItemData(cboReporte.ListIndex)
      Case "1" 'Activos y Pasivos
                  Call sbMovimientoCatalogo
      Case "2"   'Rentabilidad
                  Call sbMovimientoCatalogo
      Case "2.1" 'Rentabilidad - Centro de Costo
                  Call sbRentabilidadEspecial
      Case "2.2" 'Rentabilidad - Unidad
                  Call sbRentabilidadEspecial
      Case "3" 'Balance Completo
                  Call sbMovimientoCatalogo
      Case Else
                  Call sbMovimientoCatalogo
    End Select
     

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


cboReporte.Clear
cboReporte.AddItem "Informe de Activos y Pasivos"
cboReporte.ItemData(cboReporte.ListCount - 1) = "1"
cboReporte.AddItem "Informe de Rentabilidad"
cboReporte.ItemData(cboReporte.ListCount - 1) = "2"
cboReporte.AddItem "Rentabilidad (Centro de Costo)"
cboReporte.ItemData(cboReporte.ListCount - 1) = "2.1"
cboReporte.AddItem "Rentabilidad (Unidad)"
cboReporte.ItemData(cboReporte.ListCount - 1) = "2.2"
cboReporte.AddItem "Balance de Comprobación"
cboReporte.ItemData(cboReporte.ListCount - 1) = "3"

cboReporte.Text = "Informe de Activos y Pasivos"

End Sub
