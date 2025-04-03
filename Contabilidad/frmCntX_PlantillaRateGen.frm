VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_PlantillaRateGen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generación de Asientos - (Plantillas Porcentuales)"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4455
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   12975
      _Version        =   1310723
      _ExtentX        =   22886
      _ExtentY        =   7858
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   6855
      _Version        =   1310723
      _ExtentX        =   12091
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdMuestra 
      Height          =   315
      Left            =   9360
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Muestra"
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
   End
   Begin XtremeSuiteControls.PushButton cmdGenera 
      Height          =   315
      Left            =   10320
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Genera"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtDesPlantilla 
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Width           =   6015
      _Version        =   1310723
      _ExtentX        =   10610
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
   End
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.FlatEdit txtCodPlantilla 
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   735
      _Version        =   1310723
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
   Begin XtremeSuiteControls.FlatEdit txtDebito 
      Height          =   315
      Left            =   9000
      TabIndex        =   12
      Top             =   6600
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   7200
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCredito 
      Height          =   315
      Left            =   11040
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Débitos/Créditos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   6600
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
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
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plantilla"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1605
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmCntX_PlantillaRateGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLimpiaDatos()
 txtDesPlantilla = ""
 txtDescripcion = ""
 txtDocumento = ""
 txtMonto = "0"
 lsw.ListItems.Clear
End Sub

Private Sub cmdGenera_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, vConsecutivo As Long
Dim curDebito As Currency, curCredito As Currency
Dim vNumAsiento As String, vTipoAsiento As String


On Error GoTo vError


If Not fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
  Me.MousePointer = vbDefault
  MsgBox "El Periodo Actual se Encuentra Cerrado o no se ha creado, verifique...", vbExclamation
  Exit Sub
End If


curMonto = CCur(txtMonto)

strSQL = "select * from CntX_Plantilla_Rate" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & txtCodPlantilla
Call OpenRecordSet(rs, strSQL, 0)

'Consecutivos de Cntx_Asientos de esta Plantilla
vConsecutivo = rs!Consecutivo + 1
    
strSQL = "update CntX_Plantilla_Rate set consecutivo = " & vConsecutivo _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_plantilla = " & txtCodPlantilla
Call ConectionExecute(strSQL, 0)
    
vTipoAsiento = Trim(rs!Tipo_Asiento)
vNumAsiento = "PTR" & Format(rs!cod_plantilla, "000") & "-" & Format(vConsecutivo, "000000")
   
'Crea Maestro del Asiento
strSQL = "insert into Cntx_Asientos(cod_contabilidad,tipo_asiento,num_asiento,descripcion,fecha_asiento,balanceado,anio,mes" _
       & ",user_crea,modulo,notas) values(" & gCntX_Parametros.CodigoConta & ",'" & vTipoAsiento & "','" _
       & vNumAsiento & "','" & txtDescripcion & "','" & gCntX_Parametros.PeriodoAnio & "/" & Format(gCntX_Parametros.PeriodoMes, "00") _
       & "/01','S'," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes & ",'" & glogon.Usuario & "',20,'GENERADO CON " _
       & "PLANTILLA RATE COD : " & Format(rs!cod_plantilla, "000") & "')"
Call ConectionExecute(strSQL, 0)
     
rs.Close
   
      
strSQL = "select D.*,C.descripcion " _
       & " from CntX_Plantilla_Rate_Detalle D inner join CntX_Cuentas C on D.cod_contabilidad = C.cod_contabilidad" _
       & " and D.cod_cuenta = C.cod_cuenta" _
       & " Where D.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and D.cod_plantilla = " & txtCodPlantilla _
       & " order by D.num_linea"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
     curDebito = curMonto * (rs!Debitos / 100)
     curCredito = curMonto * (rs!Creditos / 100)
     strSQL = "insert Cntx_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,cod_cuenta," _
            & "Monto_Debito,Monto_credito,Documento,Detalle,num_linea,cod_unidad,cod_centro_costo,cod_divisa,Tipo_Cambio) values(" & gCntX_Parametros.CodigoConta _
            & ",'" & vTipoAsiento & "','" & vNumAsiento & "','" & Trim(rs!cod_cuenta) & "'," _
            & curDebito & "," & curCredito & ",'" & Trim(txtDocumento) & "','" _
            & Trim(rs!Detalle) & "'," & rs!Num_linea & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!cod_Divisa & "',1)"
     Call ConectionExecute(strSQL, 0)
 rs.MoveNext
Loop
rs.Close

MsgBox "Asiento Aplicado : " & vNumAsiento, vbInformation
Call sbLimpiaDatos

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdMuestra_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curMonto As Currency
Dim curDebitos As Currency, curCreditos As Currency

On Error GoTo vError

lsw.ListItems.Clear
curMonto = CCur(txtMonto)
curDebitos = 0
curCreditos = 0

strSQL = "select D.*,c.descripcion,U.descripcion as UniDes" _
       & " from CntX_Plantilla_Rate_Detalle D inner join CntX_Cuentas C on D.cod_contabilidad = C.cod_contabilidad" _
       & " and D.cod_cuenta = C.cod_cuenta" _
       & " inner join CntX_Unidades U on D.cod_contabilidad = U.cod_contabilidad and D.cod_unidad = U.cod_unidad" _
       & " Where D.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and D.cod_plantilla = " & txtCodPlantilla _
       & " order by D.num_linea"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , fxCntX_CuentaFormato(True, rs!cod_cuenta))
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = Format(curMonto * (rs!Debitos / 100), "Standard")
     itmX.SubItems(3) = Format(curMonto * (rs!Creditos / 100), "Standard")
     itmX.SubItems(4) = rs!Detalle
     itmX.SubItems(5) = rs!UniDes
     itmX.SubItems(6) = rs!cod_centro_costo
     itmX.SubItems(7) = rs!cod_Divisa

     curDebitos = curDebitos + (curMonto * (rs!Debitos / 100))
     curCreditos = curCreditos + (curMonto * (rs!Creditos / 100))
 
 rs.MoveNext
Loop
rs.Close

txtDebito = Format(curDebitos, "Standard")
txtCredito = Format(curCreditos, "Standard")

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
vModulo = 20
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2100
    .Add , , "Descripción", 3800
    .Add , , "Débitos", 2100, vbRightJustify
    .Add , , "Créditos", 2100, vbRightJustify
    .Add , , "Detalle", 2100
    .Add , , "Unidad", 2500, vbCenter
    .Add , , "C.C.", 2500, vbCenter
    .Add , , "Divisa", 1200, vbCenter
End With


Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub txtCodPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesPlantilla.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "cod_plantilla"
    gBusquedas.Orden = "cod_plantilla"
    gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
    gBusquedas.Consulta = "select cod_plantilla,descripcion from CntX_Plantilla_Rate"
    frmBusquedas.Show vbModal
    txtCodPlantilla = gBusquedas.Resultado
    txtDesPlantilla = gBusquedas.Resultado2
    txtCodPlantilla.SetFocus
End If

End Sub

Private Sub txtCodPlantilla_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Call sbLimpiaDatos

If Not IsNumeric(txtCodPlantilla.Text) Then Exit Sub

strSQL = "select * from CntX_Plantilla_Rate where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_plantilla = " & txtCodPlantilla.Text
Call OpenRecordSet(rs, strSQL, 0)
  txtDesPlantilla = rs!Descripcion & ""
rs.Close
vError:

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub txtDesPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
    gBusquedas.Consulta = "select cod_plantilla,descripcion from CntX_Plantilla_Rate"
    frmBusquedas.Show vbModal
    txtCodPlantilla = gBusquedas.Resultado
    txtDesPlantilla = gBusquedas.Resultado2
    txtCodPlantilla.SetFocus
End If

End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdMuestra.SetFocus
End Sub
