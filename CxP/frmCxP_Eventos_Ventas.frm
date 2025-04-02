VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxP_Eventos_Ventas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Eventos y Ferias: Informe de Ventas"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   15165
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit txtProvNombre 
      Height          =   330
      Left            =   7080
      TabIndex        =   15
      Top             =   1080
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9551
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboEvento 
      Height          =   330
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   5400
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   7080
      TabIndex        =   4
      Top             =   1440
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9551
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   12840
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmCxP_Eventos_Ventas.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   13920
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmCxP_Eventos_Ventas.frx":0700
      ImageAlignment  =   4
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5895
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   14655
      _Version        =   524288
      _ExtentX        =   25850
      _ExtentY        =   10398
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
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
      MaxCols         =   15
      ScrollBars      =   2
      SpreadDesigner  =   "frmCxP_Eventos_Ventas.frx":086A
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtProvId 
      Height          =   330
      Left            =   5400
      TabIndex        =   14
      Top             =   1080
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   12840
      TabIndex        =   17
      Top             =   1080
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   13560
      TabIndex        =   18
      Top             =   840
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Proveedor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Ventas por Concepto de Eventos/Ferias"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   13
      Top             =   240
      Width           =   6255
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Evento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fechas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   14535
      _Version        =   1441793
      _ExtentX        =   25638
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Resultados de la busqueda:"
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
   Begin VB.Image imgBanner 
      Height          =   765
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmCxP_Eventos_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    
    vHeaders.Columnas = 15
    vHeaders.Headers(1) = "Venta Id"
    vHeaders.Headers(2) = "Fecha"
    vHeaders.Headers(3) = "Vendedor"
    vHeaders.Headers(4) = "Estado"
    vHeaders.Headers(5) = "Cédula"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "SubTotal"
    vHeaders.Headers(8) = "IVA"
    vHeaders.Headers(9) = "Comisión"
    vHeaders.Headers(10) = "Mnt a Girar"
    vHeaders.Headers(11) = "No. Operación"
    vHeaders.Headers(12) = "Proveedor Id"
    vHeaders.Headers(13) = "Proveedor Nombre"
    vHeaders.Headers(14) = "Evento"
    vHeaders.Headers(15) = "Detalle Venta"

    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Ferias_" & cboEvento.Text)
 
End Sub

Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)
txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)
txtProvNombre.Text = fxSysCleanTxtInject(txtProvNombre.Text)

Dim pEvento As Long, pProvId As Long, pInicio As String, pCorte As String

If cboEvento.Text = "TODOS" Then
  pEvento = 0
Else
  pEvento = cboEvento.ItemData(cboEvento.ListIndex)
End If

If IsNumeric(txtProvId.Text) Then
    pProvId = txtProvId.Text
Else
    pProvId = 0
End If


If chkTodas.Value = vbChecked Then
    pInicio = "1900/01/01 00:00:00"
    pCorte = "2200/12/31 23:59:59"
Else
    pInicio = Format(dtpInicio.Value, "YYYY-MM-DD") & " 00:00:00"
    pCorte = Format(dtpCorte.Value, "YYYY-MM-DD") & " 23:59:59"
End If

strSQL = "exec spCxP_Eventos_Ventas " & pEvento & ", '" & pInicio & "', '" & pCorte & "', " & pProvId _
       & ", '" & txtProvNombre.Text & "', '" & txtCedula.Text & "', '" & txtNombre.Text _
       & "', '" & txtUsuario.Text & "', 'Ferias'"
       

Call OpenRecordSet(rs, strSQL)


With vGrid
  .MaxRows = 0
  Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .col = 1
     .Text = Trim(rs!ID_Venta)
     .col = 2
     .Text = Format(rs!Registro_Fecha, "yyyy-MM-dd hh:mm:ss")
     .col = 3
     .Text = Trim(rs!Registro_Usuario & "")
     .col = 4
     .Text = Trim(rs!Estado & "")
     
     .col = 5
     .Text = Trim(rs!Cliente_Cedula)
     .col = 6
     .Text = Trim(rs!Cliente_Nombre)
     
     .col = 7
     .Text = Format(rs!V_Sub_Total, "Standard")
     .col = 8
     .Text = Format(rs!V_IVA, "Standard")
     .col = 9
     .Text = Format(rs!CxP_Comision, "Standard")
     .col = 10
     .Text = Format(rs!Monto_Girar, "Standard")
     
     
     .col = 11
     .Text = Trim(CStr(rs!Crd_Operacion & ""))
     .col = 12
     .Text = Trim(rs!Cod_Proveedor)
     .col = 13
     .Text = Trim(rs!Proveedor_Desc)
     .col = 14
     .Text = Trim(rs!Evento_Desc)
     .col = 15
     .Text = Trim(rs!V_Descripcion & "")

   rs.MoveNext
  Loop
  rs.Close
End With

scMain.Caption = "Casos localizados: " & vGrid.MaxRows

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub btnAccion_Click(Index As Integer)
Select Case Index
    Case 0 'Buscar
        Call sbBuscar
        
    Case 1 'Exportar
        Call sbExportar
        
End Select
End Sub


Private Sub chkTodas_Click()
If chkTodas.Value = xtpChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub Form_Load()
vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle
vGrid.MaxRows = 0

strSQL = "Select Cod_Evento as 'IdX', Descripcion as 'ItmX' from CxP_Eventos order by Fecha_Inicio desc"
Call sbCbo_Llena_New(cboEvento, strSQL, True, True)


dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -60, dtpCorte.Value)


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 320
vGrid.Height = Me.Height - (vGrid.top + 700)

scMain.Width = vGrid.Width

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cedula"
   gBusquedas.Orden = "cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "select cedula,nombre from socios"
   frmBusquedas.Show vbModal
   txtNombre.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula.Text = Trim(gBusquedas.Resultado)
      txtNombre.Text = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

End Sub

Private Sub txtProvId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor, cedjur, descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtProvId.Text = gBusquedas.Resultado
    txtProvNombre.Text = gBusquedas.Resultado3
  End If

End If

End Sub
