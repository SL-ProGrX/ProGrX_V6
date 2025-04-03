VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_Liquidacion_BF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Liquidación de Balanza de Cobros a Favor (Seguros Cerrados)"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todos?"
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
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   7200
      TabIndex        =   2
      Top             =   1080
      Width           =   1332
      _Version        =   1441792
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmSeguros_Liquidacion_BF.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10215
      _Version        =   524288
      _ExtentX        =   18018
      _ExtentY        =   8916
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   23
      SpreadDesigner  =   "frmSeguros_Liquidacion_BF.frx":0A1E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   6864
      Width           =   10524
      _ExtentX        =   18574
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnLiquidar 
      Height          =   492
      Left            =   8520
      TabIndex        =   4
      Top             =   1080
      Width           =   1812
      _Version        =   1441792
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Liquidar"
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
      Picture         =   "frmSeguros_Liquidacion_BF.frx":17DF
   End
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   4095
      _Version        =   1441792
      _ExtentX        =   7223
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   5760
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidación de Saldo a Favor de Seguros Cerrados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Index           =   2
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   8772
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10572
   End
End
Attribute VB_Name = "frmSeguros_Liquidacion_BF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnLiquidar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String, vConcepto As String
Dim i As Long, vAseguradora As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vAseguradora = cboAseguradora.ItemData(cboAseguradora.ListIndex)

'spSeguros_PolizaDevolucion_Masiva(@Paso smallint, @Aseguradora varchar(10), @Poliza varchar(30), @Usuario varchar(30)
'                                            , @TipoDoc varchar(10), @NumDoc varchar(30), @Concepto varchar(10))

'Paso 1
strSQL = "exec spSeguros_PolizaDevolucion_Masiva 1,'" & vAseguradora & "','','" & glogon.Usuario _
       & "','','',''"
Call OpenRecordSet(rs, strSQL)
  vTipoDoc = rs!Tipo_Documento
  vNumDoc = rs!Cod_Transaccion
  vConcepto = rs!Concepto
rs.Close

'Paso 2: Procesa Devoluciones
strSQL = ""
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 1
  If vGrid.Value = vbChecked Then
        vGrid.Col = 3
        strSQL = strSQL & Space(10) & "exec spSeguros_PolizaDevolucion_Masiva 2,'" & vAseguradora & "','" & vGrid.Text & "','" & glogon.Usuario _
               & "','" & vTipoDoc & "','" & vNumDoc & "','" & vConcepto & "'"
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           strSQL = ""
        End If
  End If
Next i

'Procesa Lote pendiente
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If

'Paso 3: Asiento
strSQL = "exec spSeguros_PolizaDevolucion_Masiva 3,'" & vAseguradora & "','','" & glogon.Usuario _
       & "','" & vTipoDoc & "','" & vNumDoc & "','" & vConcepto & "'"
Call ConectionExecute(strSQL)





Me.MousePointer = vbDefault
MsgBox "Devoluciones aplicadas satisfactoriamente!", vbInformation

Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Call sbBuscar

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboAseguradora_Click()
If Not vPaso Then
  vGrid.MaxRows = 0
End If
End Sub

Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 1
  vGrid.Value = chkTodos.Value
Next i

End Sub

Private Sub Form_Activate()
 vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 17

cboEstado.Clear
cboEstado.AddItem "Cerradas"
cboEstado.AddItem "Activas"
cboEstado.Text = "Cerradas"

Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.MaxRows = 0

vPaso = True

strSQL = "select rtrim(COD_ASEGURADORA) as 'IdX', rtrim(NOMBRE) as ItmX from SEGUROS_ASEGURADORAS where Activo = 1 order by NOMBRE"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)

vPaso = False



End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 500
vGrid.Height = Me.Height - (vGrid.Top + StatusBarX.Height + 550)



End Sub



Private Sub sbBuscar()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select " & chkTodos.Value & ",COD_ASEGURADORA, NUM_POLIZA,rtrim(CEDULA),rtrim(NOMBRE),Estado_Desc" _
       & " ,CUOTA,MONTO,REGISTRO_FECHA,ACTIVA_FECHA,CIERRA_FECHA , PAGADO_TOTAL,COBRADO_TOTAL , Balanza_Cobro" _
       & " ,comision_Vendedor_Total,comision_Comercializa_Total,Comision_Interna_Total,isnull(Operacion,0), COD_PRODUCTO_Desc, Vendedor_NOMBRE" _
       & " ,Comercializadora_Nombre,Cliente_Cor_Nombre,Vendedor_Comision_Real" _
       & "  from  vSeguros_ListadoGeneral " _
       & " where Estado = '" & Mid(cboEstado.Text, 1, 1) & "' and COBRADO_TOTAL > PAGADO_TOTAL"


If cboAseguradora.Text <> "TODOS" Then
   strSQL = strSQL & " and COD_ASEGURADORA = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
End If

Call sbCargaGridLocal(vGrid, 23, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

On Error GoTo vError

vPaso = True

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i

    If rs.Fields(i - 1).Type = 135 Then
        If Year(rs.Fields(i - 1).Value) > 1900 Then
           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
        End If
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    End If
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Cobrado_Total - rs!Pagado_Total
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Balanza..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
'
'
'Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Dim vHeaders As vGridHeaders
'    vHeaders.Columnas = 23
'    vHeaders.Headers(1) = "..."
'    vHeaders.Headers(2) = "Aseguradora"
'    vHeaders.Headers(3) = "No.Póliza"
'    vHeaders.Headers(4) = "Cédula"
'    vHeaders.Headers(5) = "Nombre"
'    vHeaders.Headers(6) = "Estado"
'    vHeaders.Headers(7) = "Mensualidad"
'    vHeaders.Headers(8) = "Monto"
'    vHeaders.Headers(9) = "Fec.Registro"
'    vHeaders.Headers(10) = "Fec.Activación"
'    vHeaders.Headers(11) = "Fec.Cierre"
'    vHeaders.Headers(12) = "Total Pagado"
'    vHeaders.Headers(13) = "Total Cobrado"
'    vHeaders.Headers(14) = "Balanza Cobraza"
'    vHeaders.Headers(15) = "Comisión Vendedor"
'    vHeaders.Headers(16) = "Comisión Comercializa"
'    vHeaders.Headers(17) = "Comisión Interna"
'    vHeaders.Headers(18) = "No. Operación"
'    vHeaders.Headers(19) = "Tipo Seguro"
'    vHeaders.Headers(20) = "Vendedor"
'    vHeaders.Headers(21) = "Comercializadora"
'    vHeaders.Headers(22) = "Cliente Corporativo"
'    vHeaders.Headers(23) = "Comision Real Vendedor"
'
'Select Case ButtonMenu.Key
'  Case "Excel"
'      Call sbSIFGridExportar(vGrid, vHeaders, "SEGUROS_ConsultaPolizas")
'  Case "HTML"
'      Call sbSIFGridExportar(vGrid, vHeaders, "SEGUROS_ConsultaPolizas", "HTML")
'End Select
'End Sub
'
'
'
'
'
'
